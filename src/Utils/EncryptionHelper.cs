using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Nedev.DocToDocx.Utils;

/// <summary>
/// Decrypts XOR-encrypted streams from Word documents.
/// Implements the XOR obfuscation described in MS-DOC.
/// </summary>
public static class EncryptionHelper
{
    /// <summary>
    /// XOR decryption key derived from the document's LKey.
    /// </summary>
    private const uint DECRYPTION_KEY = 0xE1B0C1B2;

    /// <summary>
    /// Decrypts a stream using Word's XOR obfuscation.
    /// </summary>
    /// <param name="encryptedStream">The encrypted stream.</param>
    /// <param name="key">The XOR key (LKey from FIB).</param>
    /// <returns>A new stream with decrypted data.</returns>
    public static Stream DecryptXor(Stream encryptedStream, uint key)
    {
        var decryptedStream = new MemoryStream();
        
        // Read all bytes from the encrypted stream
        var buffer = new byte[4096];
        int bytesRead;
        
        while ((bytesRead = encryptedStream.Read(buffer, 0, buffer.Length)) > 0)
        {
            // XOR decrypt each byte
            for (int i = 0; i < bytesRead; i++)
            {
                buffer[i] ^= (byte)(key >> (i % 4) * 8);
            }
            
            decryptedStream.Write(buffer, 0, bytesRead);
        }
        
        decryptedStream.Position = 0;
        return decryptedStream;
    }

    /// <summary>
    /// Decrypts a byte array using Word's XOR obfuscation.
    /// </summary>
    /// <param name="encryptedBytes">The encrypted bytes.</param>
    /// <param name="key">The XOR key (LKey from FIB).</param>
    /// <returns>A new byte array with decrypted data.</returns>
    public static byte[] DecryptXor(byte[] encryptedBytes, uint key)
    {
        var decryptedBytes = new byte[encryptedBytes.Length];
        
        for (int i = 0; i < encryptedBytes.Length; i++)
        {
            decryptedBytes[i] = (byte)(encryptedBytes[i] ^ (byte)(key >> (i % 4) * 8));
        }
        
        return decryptedBytes;
    }

    /// <summary>
    /// Checks if a stream is encrypted using Word's XOR obfuscation.
    /// </summary>
    /// <param name="stream">The stream to check.</param>
    /// <param name="key">The XOR key (LKey from FIB).</param>
    /// <returns>True if the stream appears to be encrypted.</returns>
    public static bool IsXorEncrypted(Stream stream, uint key)
    {
        // Read first few bytes and check for common Word document signatures
        var buffer = new byte[1024];
        var originalPosition = stream.Position;
        
        stream.Read(buffer, 0, Math.Min(buffer.Length, (int)(stream.Length - stream.Position)));
        stream.Position = originalPosition;
        
        // Check for common Word document magic numbers
        if (buffer.Length >= 2)
        {
            var magic = (ushort)(buffer[0] | (buffer[1] << 8));
            if (magic == 0xA5EC || magic == 0xA5B3)
            {
                return false; // Not encrypted or already decrypted
            }
        }
        
        // If we can't determine, assume it might be encrypted
        return true;
    }
}