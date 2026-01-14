import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Generate an encrypted OOXML spreadsheet (`.xlsx`/`.xlsm`/`.xlsb`) by wrapping a plaintext OOXML
 * package
 * in an OLE2/CFB container ("EncryptedPackage" + "EncryptionInfo" streams).
 *
 * <p>Usage:
 *
 * <pre>
 *   java -cp ... GenerateEncryptedXlsx agile    password in.xlsx out.xlsx
 *   java -cp ... GenerateEncryptedXlsx standard password in.xlsm out.xlsm
 *   java -cp ... GenerateEncryptedXlsx agile    password in.xlsb out.xlsb
 * </pre>
 *
 * <p>Notes:
 * <ul>
 *   <li>The output file is <b>not</b> a ZIP file even if it uses a `.xlsx`/`.xlsm`/`.xlsb` extension;
 *       it is an OLE2/CFB container as used by Excel for encrypted OOXML.
 *   <li>Apache POI uses random salts/IVs for encryption, so output bytes are not expected to be
 *       bit-for-bit stable across runs. The resulting files should still be valid encrypted
 *       workbooks.</li>
 * </ul>
 */
public final class GenerateEncryptedXlsx {
  private static void usageAndExit() {
    System.err.println(
        "Usage: GenerateEncryptedXlsx <mode> <password> <in_plaintext_ooxml_zip> <out_encrypted_ooxml>\n"
            + "  mode: agile | standard\n"
            + "\n"
            + "Example:\n"
            + "  GenerateEncryptedXlsx agile password fixtures/xlsx/basic/basic.xlsx fixtures/encrypted/ooxml/agile.xlsx\n");
    System.exit(2);
  }

  private static EncryptionMode parseMode(String s) {
    if (s == null) {
      throw new IllegalArgumentException("mode is required");
    }
    switch (s.toLowerCase()) {
      case "agile":
        return EncryptionMode.agile;
      case "standard":
        return EncryptionMode.standard;
      default:
        throw new IllegalArgumentException("unknown mode: " + s + " (expected: agile|standard)");
    }
  }

  public static void main(String[] args) throws Exception {
    if (args.length != 4) {
      usageAndExit();
      return;
    }

    final EncryptionMode mode;
    try {
      mode = parseMode(args[0]);
    } catch (IllegalArgumentException e) {
      System.err.println(e.getMessage());
      usageAndExit();
      return;
    }

    final String password = args[1];
    final Path inPath = Paths.get(args[2]);
    final Path outPath = Paths.get(args[3]);

    if (!Files.isRegularFile(inPath)) {
      throw new IllegalArgumentException("input file does not exist or is not a regular file: " + inPath);
    }
    // Guard against accidentally overwriting the input (particularly when invoking the Java class directly
    // rather than via `generate.sh`, which also checks this).
    if (inPath.toAbsolutePath().normalize().equals(outPath.toAbsolutePath().normalize())) {
      throw new IllegalArgumentException("output path must be different from input path: " + inPath);
    }

    final Path parent = outPath.toAbsolutePath().getParent();
    if (parent != null) {
      Files.createDirectories(parent);
    }

    try (POIFSFileSystem fs = new POIFSFileSystem()) {
      EncryptionInfo info = new EncryptionInfo(mode);
      Encryptor enc = info.getEncryptor();
      enc.confirmPassword(password);

      // Encrypt the raw OOXML ZIP bytes directly (avoid parsing/repacking and avoid mutating the input).
      try (InputStream plaintext = Files.newInputStream(inPath);
          OutputStream encryptedStream = enc.getDataStream(fs)) {
        // OOXML packages (`.xlsx`/`.xlsm`/`.xlsb`) are ZIP files, so the payload should begin with "PK".
        // If it doesn't, still proceed
        // (the user may intentionally encrypt arbitrary bytes), but warn to avoid accidental misuse.
        byte[] first2 = new byte[2];
        int n = plaintext.read(first2);
        if (n == 2) {
          if (first2[0] != 'P' || first2[1] != 'K') {
            System.err.println("Warning: input does not look like a ZIP/OOXML package (missing PK signature).");
          }
          encryptedStream.write(first2);
        } else if (n > 0) {
          encryptedStream.write(first2, 0, n);
        }

        byte[] buf = new byte[16 * 1024];
        int read;
        while ((read = plaintext.read(buf)) != -1) {
          encryptedStream.write(buf, 0, read);
        }
      }

      try (OutputStream out = Files.newOutputStream(outPath)) {
        fs.writeFilesystem(out);
      }
    }
  }
}
