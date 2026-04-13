# DLManager

In-memory + on-disk cache for Outlook Exchange Distribution Lists, designed for VSTO Outlook add-ins.

Fetching DL members from Exchange is slow. This manager caches expanded member lists in memory and persists them to a local XML file so that subsequent Outlook sessions load instantly from disk — with no Exchange round-trips until a list goes stale.

## Requirements

- .NET Framework 4.8
- `Microsoft.Office.Interop.Outlook` (provided by the VSTO SDK)

No additional NuGet packages required.

## Setup

### 1. Copy `DistributionListManager.cs` into your VSTO project

The file targets the `DLManager` namespace. Change it to match your project's namespace if needed.

### 2. Wire it up in `ThisAddIn.cs`

Expose the manager as a public property so any class in the add-in can reach it via `Globals.ThisAddIn.DLManager`.

```csharp
// ThisAddIn.cs
public partial class ThisAddIn
{
    public DistributionListManager DLManager { get; private set; }

    private void ThisAddIn_Startup(object sender, EventArgs e)
    {
        // Must be constructed on the Outlook STA thread — ThisAddIn_Startup qualifies.
        DLManager = new DistributionListManager(this.Application);

        // Phase 1: loads on-disk cache instantly (no Exchange traffic).
        // Phase 2: syncs new top-level DL stubs from Exchange in the background.
        DLManager.BeginPreload();
    }

    private void ThisAddIn_Shutdown(object sender, EventArgs e)
    {
        DLManager?.Dispose(); // flushes any pending file save before exit
        DLManager = null;
    }
}
```

## Usage

### Get all members of a DL (by SMTP address)

```csharp
var members = await Globals.ThisAddIn.DLManager.GetMembersBySmtpAsync("engineering@contoso.com");

foreach (var m in members)
    Console.WriteLine($"{m.DisplayName} <{m.SmtpAddress}>");
```

### Get all members of a DL (by EntryID)

```csharp
// entryId comes from AddressEntry.ID, a recipient's AddressEntry, etc.
var members = await Globals.ThisAddIn.DLManager.GetMembersAsync(entryId);
```

Both methods return a flat `IReadOnlyList<MemberInfo>`. Nested DLs are expanded recursively — only leaf members (mailboxes/contacts) are included.

### Browse preloaded top-level DL stubs

```csharp
// Available immediately after BeginPreload() finishes — no members, just metadata.
var dls = Globals.ThisAddIn.DLManager.GetTopLevelDLs();

foreach (var (name, smtp, id) in dls)
    Console.WriteLine($"{name} <{smtp}>");
```

### Invalidate a stale list

```csharp
// Force the next GetMembersAsync call to re-fetch from Exchange.
Globals.ThisAddIn.DLManager.Invalidate(entryId);

// Or clear all cached member lists at once.
Globals.ThisAddIn.DLManager.InvalidateAllMembers();
```

## Configuration

All parameters are optional — defaults are shown below.

```csharp
DLManager = new DistributionListManager(
    app:           this.Application,
    ttl:           TimeSpan.FromMinutes(30),       // how long a member list stays fresh
    cacheFilePath: @"C:\custom\path\dlcache.xml"  // defaults to %LOCALAPPDATA%\DLManager\dlcache.xml
);
```

### Preload options

```csharp
DLManager.BeginPreload(new DistributionListManager.PreloadOptions
{
    // Scan a specific Exchange address list instead of the entire GAL.
    // Useful for large organisations — keeps startup fast.
    AddressListName = "All Distribution Lists",

    // Cap the number of top-level DL stubs loaded from Exchange (0 = no limit).
    MaxEntries = 500,

    // Set false to skip the Exchange sync and rely entirely on the on-disk cache.
    RefreshFromExchangeOnStartup = true,
});
```

## How it works

```
Outlook startup
    │
    ├─ BeginPreload()
    │       │
    │       ├─ Phase 1: Read dlcache.xml → populate in-memory cache (instant)
    │       │
    │       └─ Phase 2: Sync top-level DL stubs from Exchange (background, STA thread)
    │                   TryAdd only — does not overwrite entries loaded from file
    │
    └─ User triggers GetMembersBySmtpAsync("eng@contoso.com")
            │
            ├─ Cache hit (fresh): return immediately from memory
            │
            └─ Cache miss / stale: expand from Exchange (STA thread, non-blocking)
                    │
                    ├─ Recursively expand nested DLs (cycle-safe)
                    ├─ Write result into in-memory cache
                    └─ Schedule debounced save → dlcache.xml (3 s after last write)
```

### Cache file

Location: `%LOCALAPPDATA%\DLManager\dlcache.xml`

Only DLs whose member lists have been expanded are written to the file. DLs the user has never accessed are never fetched or stored. The file is written via a `.tmp`-then-swap pattern so it is never left half-written if Outlook exits unexpectedly.

Example file:

```xml
<?xml version="1.0" encoding="utf-8"?>
<dlCache version="1" savedAt="2024-01-15T10:30:00.0000000Z">
  <entry entryId="AAAAAA..." displayName="Engineering" smtpAddress="eng@contoso.com"
         fetchedAt="2024-01-15T10:00:00.0000000Z">
    <member displayName="Alice Smith" smtpAddress="alice@contoso.com"
            entryId="BBBBBB..." isDL="false" />
    <member displayName="Bob Jones" smtpAddress="bob@contoso.com"
            entryId="CCCCCC..." isDL="false" />
  </entry>
</dlCache>
```

## Threading notes

- `DistributionListManager` **must be constructed on the Outlook STA thread** (`ThisAddIn_Startup` satisfies this). It captures the `SynchronizationContext` at construction time and uses it to marshal all Outlook COM calls back to the STA thread.
- All public methods (`GetMembersAsync`, `GetMembersBySmtpAsync`, etc.) are safe to call from any thread.
- File I/O runs on the thread pool and never touches the STA thread.
- Concurrent callers requesting the same DL share one in-flight fetch — Exchange is contacted only once per DL regardless of how many callers are waiting.
