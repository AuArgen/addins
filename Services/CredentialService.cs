using System.Runtime.InteropServices;
using System.Text;

namespace AIDocAssistant.Services;

// Хранение API-ключа через Windows Credential Manager
public static class CredentialService
{
    private const string TargetPrefix = "AIDocAssistant_";

    public static void Save(string provider, string apiKey)
    {
        var target      = TargetPrefix + provider;
        var credBytes   = Encoding.Unicode.GetBytes(apiKey);
        var credential  = new CREDENTIAL
        {
            Type                 = 1, // CRED_TYPE_GENERIC
            TargetName           = target,
            CredentialBlobSize   = (uint)credBytes.Length,
            CredentialBlob       = Marshal.AllocHGlobal(credBytes.Length),
            Persist              = 2, // CRED_PERSIST_LOCAL_MACHINE
            UserName             = Environment.UserName
        };
        Marshal.Copy(credBytes, 0, credential.CredentialBlob, credBytes.Length);

        try { CredWrite(ref credential, 0); }
        finally { Marshal.FreeHGlobal(credential.CredentialBlob); }
    }

    public static string? Load(string provider)
    {
        var target = TargetPrefix + provider;
        if (!CredRead(target, 1, 0, out var credPtr))
            return null;

        try
        {
            var cred  = Marshal.PtrToStructure<CREDENTIAL>(credPtr);
            var bytes = new byte[cred.CredentialBlobSize];
            Marshal.Copy(cred.CredentialBlob, bytes, 0, bytes.Length);
            return Encoding.Unicode.GetString(bytes);
        }
        finally { CredFree(credPtr); }
    }

    public static void Delete(string provider)
    {
        CredDelete(TargetPrefix + provider, 1, 0);
    }

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredWrite(ref CREDENTIAL credential, uint flags);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern void CredFree(IntPtr buffer);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredDelete(string target, uint type, uint flags);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct CREDENTIAL
    {
        public uint    Flags;
        public uint    Type;
        public string  TargetName;
        public string? Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint    CredentialBlobSize;
        public IntPtr  CredentialBlob;
        public uint    Persist;
        public uint    AttributeCount;
        public IntPtr  Attributes;
        public string? TargetAlias;
        public string  UserName;
    }
}
