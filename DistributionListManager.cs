using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DLManager
{
    /// <summary>
    /// Cache for Outlook Exchange Distribution Lists backed by a local XML file.
    ///
    /// Threading model:
    ///   Outlook COM objects are apartment-threaded (STA). This class captures the STA
    ///   SynchronizationContext at construction and marshals every COM call back to it.
    ///   All public methods are safe to call from any thread and return Tasks.
    ///
    /// Startup strategy:
    ///   Call <see cref="BeginPreload"/> once from ThisAddIn_Startup.
    ///   It first loads the on-disk cache (no Exchange round-trips), then syncs new
    ///   top-level DL stubs from Exchange in the background. Member lists for DLs the
    ///   user has never accessed are never fetched.
    ///
    /// Persistence:
    ///   After each member-list expansion, a debounced save writes the cache to
    ///   <see cref="DefaultCacheFilePath"/> (or the path you supply). Stale entries
    ///   in the file are refreshed lazily on the next <see cref="GetMembersAsync"/> call.
    ///   Writes use a temp-file swap so the cache is never left half-written.
    /// </summary>
    public sealed class DistributionListManager : IDisposable
    {
        // -------------------------------------------------------------------------
        // Public types
        // -------------------------------------------------------------------------

        /// <summary>One resolved leaf member (non-DL) returned from a recursive expansion.</summary>
        public sealed class MemberInfo
        {
            public string DisplayName        { get; }
            public string SmtpAddress        { get; }
            public string EntryId            { get; }
            public bool   IsDistributionList { get; }

            internal MemberInfo(string displayName, string smtpAddress, string entryId, bool isDl)
            {
                DisplayName        = displayName;
                SmtpAddress        = smtpAddress;
                EntryId            = entryId;
                IsDistributionList = isDl;
            }
        }

        /// <summary>Options that control which address list is scanned at startup.</summary>
        public sealed class PreloadOptions
        {
            /// <summary>
            /// Name of the Exchange address list to scan for top-level DLs.
            /// Leave null to scan the Global Address List.
            /// Example: "All Distribution Lists"
            /// </summary>
            public string AddressListName { get; set; }

            /// <summary>
            /// Maximum number of top-level DL stubs to preload from Exchange.
            /// 0 (default) means no limit.
            /// </summary>
            public int MaxEntries { get; set; } = 0;

            /// <summary>
            /// When true (default), after loading from the on-disk cache, a background
            /// Exchange sync runs to pick up any DLs added since the last save.
            /// </summary>
            public bool RefreshFromExchangeOnStartup { get; set; } = true;
        }

        // -------------------------------------------------------------------------
        // Private types — in-memory cache entry
        // -------------------------------------------------------------------------

        private sealed class CacheEntry
        {
            public string EntryId     { get; }
            public string DisplayName { get; }
            public string SmtpAddress { get; }

            // null  → metadata cached but members not yet expanded
            // !null → last successful expansion result
            public IReadOnlyList<MemberInfo> Members;
            public DateTime MembersFetchedAt;

            public CacheEntry(string entryId, string displayName, string smtpAddress)
            {
                EntryId     = entryId;
                DisplayName = displayName;
                SmtpAddress = smtpAddress;
            }

            public bool AreMembersStale(TimeSpan ttl) =>
                Members == null || (DateTime.UtcNow - MembersFetchedAt) > ttl;
        }

        // -------------------------------------------------------------------------
        // Constants & static serialization state
        // -------------------------------------------------------------------------

        private const int CacheFileVersion = 1;
        private const int SaveDebounceMs   = 3_000; // write 3 s after the last change

        // XmlSerializer is expensive to construct — one static instance per root type.
        private static readonly XmlSerializer Serializer = new(typeof(CacheFileDto));

        private static readonly XmlWriterSettings WriterSettings = new()
        {
            Indent      = true,
            Encoding    = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), // UTF-8, no BOM
            CloseOutput = false,
        };

        // -------------------------------------------------------------------------
        // Fields
        // -------------------------------------------------------------------------

        private readonly Outlook.Application    _app;
        private readonly TimeSpan               _ttl;
        private readonly SynchronizationContext _staContext;
        private readonly string                 _cacheFilePath;

        // EntryID → CacheEntry (primary cache)
        private readonly ConcurrentDictionary<string, CacheEntry> _cache =
            new ConcurrentDictionary<string, CacheEntry>(StringComparer.OrdinalIgnoreCase);

        // SMTP address → EntryID (secondary lookup index)
        private readonly ConcurrentDictionary<string, string> _smtpIndex =
            new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // EntryID → Lazy<Task> — guarantees exactly one in-flight fetch per DL
        private readonly ConcurrentDictionary<string, Lazy<Task<IReadOnlyList<MemberInfo>>>> _inFlight =
            new ConcurrentDictionary<string, Lazy<Task<IReadOnlyList<MemberInfo>>>>(StringComparer.OrdinalIgnoreCase);

        // File I/O
        private readonly SemaphoreSlim _fileLock  = new(1, 1);
        private readonly Timer         _saveTimer;
        private volatile bool          _pendingSave;

        private int _disposed; // 0 = alive, 1 = disposed

        // -------------------------------------------------------------------------
        // Constructor
        // -------------------------------------------------------------------------

        /// <summary>
        /// Creates the manager.
        /// Must be called on the Outlook STA thread (e.g. from ThisAddIn_Startup)
        /// so the STA SynchronizationContext can be captured.
        /// </summary>
        /// <param name="app">The running Outlook application instance.</param>
        /// <param name="ttl">
        /// How long a cached member list stays valid. Defaults to 30 minutes.
        /// </param>
        /// <param name="cacheFilePath">
        /// Path to the XML cache file. Defaults to <see cref="DefaultCacheFilePath"/>.
        /// </param>
        public DistributionListManager(
            Outlook.Application app,
            TimeSpan?           ttl           = null,
            string              cacheFilePath = null)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _ttl = ttl ?? TimeSpan.FromMinutes(30);
            _staContext = SynchronizationContext.Current
                ?? throw new InvalidOperationException(
                    $"{nameof(DistributionListManager)} must be constructed on the Outlook STA thread.");

            _cacheFilePath = cacheFilePath ?? DefaultCacheFilePath;

            // Inactive until ScheduleSave() arms it.
            _saveTimer = new Timer(OnSaveTimerElapsed, null, Timeout.Infinite, Timeout.Infinite);
        }

        /// <summary>Default cache file path: %LOCALAPPDATA%\DLManager\dlcache.xml</summary>
        public static string DefaultCacheFilePath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DLManager",
            "dlcache.xml");

        // -------------------------------------------------------------------------
        // Public API — preload
        // -------------------------------------------------------------------------

        /// <summary>
        /// Begins preloading (fire-and-forget, non-blocking).
        /// Phase 1: reads the on-disk cache immediately — no Exchange traffic.
        /// Phase 2: syncs new top-level DL stubs from Exchange in the background
        ///          (controlled by <see cref="PreloadOptions.RefreshFromExchangeOnStartup"/>).
        /// Call once from ThisAddIn_Startup.
        /// </summary>
        public void BeginPreload(PreloadOptions options = null)
        {
            ThrowIfDisposed();
            _ = RunPreloadAsync(options ?? new PreloadOptions());
        }

        /// <summary>Awaitable variant of <see cref="BeginPreload"/>.</summary>
        public Task PreloadAsync(PreloadOptions options = null)
        {
            ThrowIfDisposed();
            return RunPreloadAsync(options ?? new PreloadOptions());
        }

        // -------------------------------------------------------------------------
        // Public API — member retrieval
        // -------------------------------------------------------------------------

        /// <summary>
        /// Returns the flat, recursively-expanded leaf-member list for the DL identified
        /// by <paramref name="entryId"/>. Served from cache when fresh; fetched from
        /// Exchange and persisted to disk otherwise.
        /// </summary>
        public Task<IReadOnlyList<MemberInfo>> GetMembersAsync(string entryId)
        {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(entryId))
                throw new ArgumentException("EntryID must not be empty.", nameof(entryId));

            return GetOrFetchMembersAsync(entryId);
        }

        /// <summary>
        /// Resolves a DL by SMTP address, then returns its flat expanded member list.
        /// </summary>
        public async Task<IReadOnlyList<MemberInfo>> GetMembersBySmtpAsync(string smtpAddress)
        {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(smtpAddress))
                throw new ArgumentException("SMTP address must not be empty.", nameof(smtpAddress));

            if (!_smtpIndex.TryGetValue(smtpAddress, out string entryId))
            {
                entryId = await RunOnStaAsync(() => ResolveEntryIdBySmtp(smtpAddress))
                              .ConfigureAwait(false);

                if (entryId == null)
                    throw new KeyNotFoundException(
                        $"No Exchange distribution list found for SMTP address '{smtpAddress}'.");
            }

            return await GetOrFetchMembersAsync(entryId).ConfigureAwait(false);
        }

        // -------------------------------------------------------------------------
        // Public API — inspection & invalidation
        // -------------------------------------------------------------------------

        /// <summary>
        /// Returns the cached DL stubs (display name + SMTP + EntryID).
        /// Does not include member lists; call <see cref="GetMembersAsync"/> to expand them.
        /// </summary>
        public IReadOnlyList<(string DisplayName, string SmtpAddress, string EntryId)> GetTopLevelDLs()
        {
            ThrowIfDisposed();
            return _cache.Values
                         .Select(e => (e.DisplayName, e.SmtpAddress, e.EntryId))
                         .ToList();
        }

        /// <summary>
        /// Marks a single DL's member list as stale in both memory and the on-disk cache.
        /// The next <see cref="GetMembersAsync"/> call for this DL will re-fetch from Exchange.
        /// </summary>
        public void Invalidate(string entryId)
        {
            ThrowIfDisposed();
            if (_cache.TryGetValue(entryId, out var entry))
            {
                entry.Members          = null;
                entry.MembersFetchedAt = default;
                ScheduleSave();
            }
        }

        /// <summary>
        /// Marks every cached member list as stale. The on-disk cache is updated accordingly.
        /// </summary>
        public void InvalidateAllMembers()
        {
            ThrowIfDisposed();
            foreach (var entry in _cache.Values)
            {
                entry.Members          = null;
                entry.MembersFetchedAt = default;
            }
            ScheduleSave();
        }

        // -------------------------------------------------------------------------
        // Preload orchestration
        // -------------------------------------------------------------------------

        private async Task RunPreloadAsync(PreloadOptions opts)
        {
            try
            {
                // Phase 1: restore from file — instant, no Exchange traffic.
                await LoadFromFileAsync().ConfigureAwait(false);

                // Phase 2: pick up DLs added since the last save.
                if (opts.RefreshFromExchangeOnStartup)
                    await RunOnStaAsync(() => PreloadFromExchange(opts)).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DLManager] Preload failed: {ex.Message}");
            }
        }

        private void PreloadFromExchange(PreloadOptions opts)
        {
            Outlook.AddressList list = null;
            try
            {
                list = string.IsNullOrEmpty(opts.AddressListName)
                    ? _app.Session.GetGlobalAddressList()
                    : FindAddressList(opts.AddressListName);

                if (list == null)
                {
                    Debug.WriteLine("[DLManager] Exchange preload: address list not found.");
                    return;
                }

                int indexed = 0;
                foreach (Outlook.AddressEntry entry in list.AddressEntries)
                {
                    try
                    {
                        if (entry.AddressEntryUserType !=
                            Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                            continue;

                        TryCacheMetadata(entry); // TryAdd — won't overwrite file-loaded entries
                        indexed++;

                        if (opts.MaxEntries > 0 && indexed >= opts.MaxEntries)
                            break;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[DLManager] Exchange preload: skipping '{entry?.Name}': {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComSafe(entry);
                    }
                }

                Debug.WriteLine($"[DLManager] Exchange preload: {indexed} top-level DL(s) indexed.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DLManager] Exchange preload error: {ex.Message}");
            }
            finally
            {
                ReleaseComSafe(list);
            }
        }

        private Outlook.AddressList FindAddressList(string name)
        {
            foreach (Outlook.AddressList al in _app.Session.AddressLists)
            {
                if (string.Equals(al.Name, name, StringComparison.OrdinalIgnoreCase))
                    return al;
            }
            return null;
        }

        // -------------------------------------------------------------------------
        // File I/O
        // -------------------------------------------------------------------------

        private async Task LoadFromFileAsync()
        {
            if (!File.Exists(_cacheFilePath))
                return;

            await _fileLock.WaitAsync().ConfigureAwait(false);
            try
            {
                byte[] xmlBytes;
                try
                {
                    xmlBytes = await ReadBytesAsync(_cacheFilePath).ConfigureAwait(false);
                }
                catch (IOException ex)
                {
                    Debug.WriteLine($"[DLManager] Cannot read cache file: {ex.Message}");
                    return;
                }

                CacheFileDto file;
                try
                {
                    using var ms = new MemoryStream(xmlBytes);
                    file = (CacheFileDto)Serializer.Deserialize(ms);
                }
                catch (Exception ex) // InvalidOperationException wraps XmlException on bad XML
                {
                    Debug.WriteLine($"[DLManager] Cache file corrupted; discarding. ({ex.Message})");
                    return;
                }

                if (file?.Version != CacheFileVersion || file.Entries == null)
                {
                    Debug.WriteLine("[DLManager] Cache file version mismatch; discarding.");
                    return;
                }

                int loaded = 0;
                foreach (var dto in file.Entries)
                {
                    if (string.IsNullOrEmpty(dto.EntryId))
                        continue;

                    var entry = new CacheEntry(
                        dto.EntryId,
                        dto.DisplayName ?? string.Empty,
                        dto.SmtpAddress ?? string.Empty);

                    if (dto.Members != null && !string.IsNullOrEmpty(dto.FetchedAt))
                    {
                        entry.Members =
                        [
                            .. dto.Members.Select(m => new MemberInfo(
                                m.DisplayName ?? string.Empty,
                                m.SmtpAddress ?? string.Empty,
                                m.EntryId     ?? string.Empty,
                                m.IsDistributionList)),
                        ];

                        entry.MembersFetchedAt = DateTime.Parse(
                            dto.FetchedAt,
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.RoundtripKind);
                    }

                    if (_cache.TryAdd(dto.EntryId, entry) && !string.IsNullOrEmpty(dto.SmtpAddress))
                        _smtpIndex[dto.SmtpAddress] = dto.EntryId;

                    loaded++;
                }

                Debug.WriteLine($"[DLManager] Loaded {loaded} entr(ies) from '{_cacheFilePath}'.");
            }
            finally
            {
                _fileLock.Release();
            }
        }

        private async Task SaveToDiskAsync()
        {
            if (Volatile.Read(ref _disposed) == 1)
                return;

            _pendingSave = false;

            // Snapshot the cache — only persist entries with expanded members.
            var dtos = _cache.Values
                .Where(e => e.Members != null)
                .Select(e => new CacheEntryDto
                {
                    EntryId     = e.EntryId,
                    DisplayName = e.DisplayName,
                    SmtpAddress = e.SmtpAddress,
                    FetchedAt   = e.MembersFetchedAt == default
                                      ? null
                                      : e.MembersFetchedAt.ToString("O", CultureInfo.InvariantCulture),
                    Members     =
                    [
                        .. e.Members.Select(m => new MemberDto
                        {
                            DisplayName        = m.DisplayName,
                            SmtpAddress        = m.SmtpAddress,
                            EntryId            = m.EntryId,
                            IsDistributionList = m.IsDistributionList,
                        }),
                    ],
                })
                .ToList();

            var fileDto = new CacheFileDto
            {
                Version = CacheFileVersion,
                SavedAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
                Entries = dtos,
            };

            // XmlSerializer is synchronous — serialize into a MemoryStream first,
            // then write the bytes to disk asynchronously.
            byte[] xmlBytes;
            using (var ms = new MemoryStream())
            {
                using (var writer = XmlWriter.Create(ms, WriterSettings))
                    Serializer.Serialize(writer, fileDto);
                xmlBytes = ms.ToArray();
            }

            await _fileLock.WaitAsync().ConfigureAwait(false);
            try
            {
                EnsureDirectory(_cacheFilePath);

                // Write to a temp file first, then swap atomically.
                // If the process is killed mid-write the original file stays intact.
                string tmp = _cacheFilePath + ".tmp";
                await WriteBytesAsync(tmp, xmlBytes).ConfigureAwait(false);

                if (File.Exists(_cacheFilePath))
                    File.Replace(tmp, _cacheFilePath, null);
                else
                    File.Move(tmp, _cacheFilePath);

                Debug.WriteLine($"[DLManager] Saved {dtos.Count} entr(ies) to '{_cacheFilePath}'.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DLManager] Failed to save cache: {ex.Message}");
            }
            finally
            {
                _fileLock.Release();
            }
        }

        private void ScheduleSave()
        {
            _pendingSave = true;
            // Reset the countdown each call; fires SaveDebounceMs after the last change.
            _saveTimer.Change(SaveDebounceMs, Timeout.Infinite);
        }

        private void OnSaveTimerElapsed(object state)
        {
            if (Volatile.Read(ref _disposed) == 1) return;
            _ = SaveToDiskAsync();
        }

        // -------------------------------------------------------------------------
        // Fetch & expand
        // -------------------------------------------------------------------------

        private Task<IReadOnlyList<MemberInfo>> GetOrFetchMembersAsync(string entryId)
        {
            if (_cache.TryGetValue(entryId, out var cached) && !cached.AreMembersStale(_ttl))
                return Task.FromResult(cached.Members);

            // Deduplicate: if a fetch is already in-flight for this entryId, share its Task.
            var lazy = _inFlight.GetOrAdd(
                entryId,
                id => new Lazy<Task<IReadOnlyList<MemberInfo>>>(
                    () => FetchAndCacheAsync(id),
                    LazyThreadSafetyMode.ExecutionAndPublication));

            return lazy.Value;
        }

        private async Task<IReadOnlyList<MemberInfo>> FetchAndCacheAsync(string entryId)
        {
            try
            {
                var members = await RunOnStaAsync(() =>
                {
                    var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    return ExpandDlRecursive(entryId, visited);
                }).ConfigureAwait(false);

                var entry = _cache.GetOrAdd(
                    entryId,
                    id => new CacheEntry(id, string.Empty, string.Empty));

                entry.Members          = members;
                entry.MembersFetchedAt = DateTime.UtcNow;

                // Persist asynchronously; don't block the caller.
                ScheduleSave();

                return members;
            }
            finally
            {
                _inFlight.TryRemove(entryId, out _);
            }
        }

        /// <summary>
        /// Recursively expands <paramref name="entryId"/> into a flat list of leaf members.
        /// Nested DLs are traversed; only individual members are emitted.
        /// Must be called on the STA thread.
        /// </summary>
        private List<MemberInfo> ExpandDlRecursive(string entryId, HashSet<string> visited)
        {
            if (!visited.Add(entryId))
            {
                Debug.WriteLine($"[DLManager] Cycle detected at EntryID={entryId}; stopping recursion.");
                return new List<MemberInfo>();
            }

            Outlook.AddressEntry   dlEntry = null;
            Outlook.AddressEntries members = null;
            try
            {
                dlEntry = _app.Session.GetAddressEntryFromID(entryId);
                members = dlEntry?.Members;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DLManager] Cannot expand EntryID={entryId}: {ex.Message}");
                return new List<MemberInfo>();
            }

            if (members == null)
                return new List<MemberInfo>();

            var result = new List<MemberInfo>();

            foreach (Outlook.AddressEntry member in members)
            {
                try
                {
                    bool isDl = member.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry;

                    if (isDl)
                    {
                        TryCacheMetadata(member);
                        result.AddRange(ExpandDlRecursive(member.ID, visited));
                    }
                    else
                    {
                        result.Add(new MemberInfo(
                            displayName: member.Name,
                            smtpAddress: ResolveSmtp(member),
                            entryId:     member.ID,
                            isDl:        false));
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[DLManager] Skipping member '{member?.Name}': {ex.Message}");
                }
                finally
                {
                    ReleaseComSafe(member);
                }
            }

            return result;
        }

        // -------------------------------------------------------------------------
        // Helpers
        // -------------------------------------------------------------------------

        private void TryCacheMetadata(Outlook.AddressEntry entry)
        {
            Outlook.ExchangeDistributionList dl = null;
            try
            {
                dl = entry.GetExchangeDistributionList();
                if (dl == null) return;

                var id   = entry.ID;
                var name = entry.Name ?? string.Empty;
                var smtp = dl.PrimarySmtpAddress ?? string.Empty;

                _cache.TryAdd(id, new CacheEntry(id, name, smtp));

                if (smtp.Length > 0)
                    _smtpIndex[smtp] = id;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DLManager] TryCacheMetadata failed for '{entry?.Name}': {ex.Message}");
            }
            finally
            {
                ReleaseComSafe(dl);
            }
        }

        private static string ResolveSmtp(Outlook.AddressEntry entry)
        {
            Outlook.ExchangeUser exUser = null;
            try
            {
                if (entry.AddressEntryUserType ==
                    Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                {
                    exUser = entry.GetExchangeUser();
                    if (exUser != null)
                        return exUser.PrimarySmtpAddress ?? entry.Address ?? string.Empty;
                }
                return entry.Address ?? string.Empty;
            }
            catch
            {
                return entry.Address ?? string.Empty;
            }
            finally
            {
                ReleaseComSafe(exUser);
            }
        }

        private string ResolveEntryIdBySmtp(string smtpAddress)
        {
            Outlook.Recipient recip = null;
            try
            {
                recip = _app.Session.CreateRecipient(smtpAddress);
                if (!recip.Resolve()) return null;

                var ae = recip.AddressEntry;
                if (ae == null) return null;
                if (ae.AddressEntryUserType !=
                    Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                    return null;

                TryCacheMetadata(ae);
                return ae.ID;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DLManager] ResolveEntryIdBySmtp('{smtpAddress}'): {ex.Message}");
                return null;
            }
            finally
            {
                ReleaseComSafe(recip);
            }
        }

        // -------------------------------------------------------------------------
        // STA dispatch
        // -------------------------------------------------------------------------

        private Task<T> RunOnStaAsync<T>(Func<T> func)
        {
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            _staContext.Post(_ =>
            {
                try   { tcs.SetResult(func()); }
                catch (Exception ex) { tcs.SetException(ex); }
            }, state: null);
            return tcs.Task;
        }

        private Task RunOnStaAsync(Action action) =>
            RunOnStaAsync(() => { action(); return true; });

        // -------------------------------------------------------------------------
        // Async file helpers
        // -------------------------------------------------------------------------

        private static async Task<byte[]> ReadBytesAsync(string path)
        {
            using var fs = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
            var buffer = new byte[(int)fs.Length];
            await fs.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
            return buffer;
        }

        private static async Task WriteBytesAsync(string path, byte[] bytes)
        {
            using var fs = new FileStream(
                path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
            await fs.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
        }

        private static void EnsureDirectory(string filePath)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        // -------------------------------------------------------------------------
        // COM cleanup
        // -------------------------------------------------------------------------

        private static void ReleaseComSafe(object obj)
        {
            if (obj != null)
                try { Marshal.ReleaseComObject(obj); } catch { }
        }

        // -------------------------------------------------------------------------
        // Guard & IDisposable
        // -------------------------------------------------------------------------

        private void ThrowIfDisposed()
        {
            if (Volatile.Read(ref _disposed) == 1)
                throw new ObjectDisposedException(nameof(DistributionListManager));
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;

            // Disarm the debounce timer before flushing so it can't fire concurrently.
            _saveTimer.Change(Timeout.Infinite, Timeout.Infinite);
            _saveTimer.Dispose();

            // Flush any pending save synchronously so data isn't lost on shutdown.
            if (_pendingSave)
                SaveToDiskAsync().GetAwaiter().GetResult();

            _fileLock.Dispose();
            _cache.Clear();
            _smtpIndex.Clear();
            _inFlight.Clear();
        }
    }

    // =========================================================================
    // XML serialization DTOs
    //
    // Kept internal (not nested) because XmlSerializer requires the types it
    // serializes to be accessible — private nested classes are not.
    //
    // Produced XML structure:
    //
    //   <?xml version="1.0" encoding="utf-8"?>
    //   <dlCache version="1" savedAt="2024-01-15T10:30:00.0000000Z">
    //     <entry entryId="..." displayName="Engineering"
    //            smtpAddress="eng@contoso.com"
    //            fetchedAt="2024-01-15T10:00:00.0000000Z">
    //       <member displayName="Alice Smith" smtpAddress="alice@contoso.com"
    //               entryId="..." isDL="false" />
    //     </entry>
    //   </dlCache>
    // =========================================================================

    [XmlRoot("dlCache")]
    internal sealed class CacheFileDto
    {
        [XmlAttribute("version")] public int                  Version { get; set; }
        [XmlAttribute("savedAt")] public string               SavedAt { get; set; }
        [XmlElement("entry")]     public List<CacheEntryDto>  Entries { get; set; }
    }

    internal sealed class CacheEntryDto
    {
        [XmlAttribute("entryId")]     public string         EntryId     { get; set; }
        [XmlAttribute("displayName")] public string         DisplayName { get; set; }
        [XmlAttribute("smtpAddress")] public string         SmtpAddress { get; set; }
        /// <summary>ISO 8601 UTC timestamp; null/empty when members have not been expanded.</summary>
        [XmlAttribute("fetchedAt")]   public string         FetchedAt   { get; set; }
        [XmlElement("member")]        public List<MemberDto> Members     { get; set; }
    }

    internal sealed class MemberDto
    {
        [XmlAttribute("displayName")] public string DisplayName        { get; set; }
        [XmlAttribute("smtpAddress")] public string SmtpAddress        { get; set; }
        [XmlAttribute("entryId")]     public string EntryId            { get; set; }
        [XmlAttribute("isDL")]        public bool   IsDistributionList { get; set; }
    }
}
