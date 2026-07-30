# DLManager

## Requirements

- .NET 8 SDK

## Build & run

```bash
dotnet build
dotnet run
```

## Creating test cases
When creating test cases, it is recommended and suggested to use the following format:

```csharp
[TestClass]
public class SomethingTests
{
    [TestMethod]
    public void MethodName_Scenario_ExpectedBehavior()
    {
        // Arrange - set up inputs/objects
        // Act    - call the method under test
        // Assert - Assert.AreEqual / IsTrue / IsNull / ThrowsException, etc.
    }
}
```

Name the test case like so: `MethodUnderTest_Scenario_ExpectedResult`