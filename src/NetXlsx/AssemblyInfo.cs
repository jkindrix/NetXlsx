using System.Runtime.CompilerServices;

// Allow the test project to see internal types via [InternalsVisibleTo].
// Strong-name signing requires the PublicKey of the test assembly. Both
// assemblies are signed with the same netxlsx.snk, so the public key
// matches. The public key is committed below; it is *not* secret (only the
// private key is).
//
// To regenerate this attribute's PublicKey after rotating the SNK:
//   sn -p netxlsx.snk netxlsx.pub
//   sn -tp netxlsx.pub   # prints the public-key blob
//
// Equivalent on Unix (requires the .NET strong-name tool from the SDK):
//   dotnet tool install --global dotnet-strong-name
//
// Until the InternalsVisibleTo public key is regenerated alongside the SNK,
// this attribute is commented out. Add it back when wiring the test project's
// internals visibility need.
//
// [assembly: InternalsVisibleTo("NetXlsx.Tests, PublicKey=...")]
