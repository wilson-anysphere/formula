/// Human-readable `VBA/dir` record names for debugging.
///
/// Record IDs are from **MS-OVBA §2.3.4 (dir Stream)**.
///
/// Notes (seen in the wild):
/// - Some `VBA/dir` encodings emit Unicode/alternate strings as standalone records whose `data` is
///   UTF-16LE bytes (sometimes with an internal `u32le` length prefix).
/// - The exact record IDs used for these Unicode/alternate records vary by layout; in this repo we
///   most commonly see:
///   - Project strings: `0x0040` (docstring Unicode), `0x003D` (helpfilepath2), `0x003C` (constants Unicode)
///   - Module strings: `0x0047` (modulename Unicode), `0x0032` (modulestreamname Unicode), `0x0048` (moduledocstring Unicode)
/// - `0x004A` is `PROJECTCOMPATVERSION` (present in many real-world files but excluded from some
///   hashing transcripts).
pub fn record_name(id: u16) -> Option<&'static str> {
    Some(match id {
        // ---- Project information records ----
        0x0001 => "PROJECTSYSKIND",
        0x0002 => "PROJECTLCID",
        0x0003 => "PROJECTCODEPAGE",
        0x0004 => "PROJECTNAME",
        0x0005 => "PROJECTDOCSTRING",
        0x0040 => "PROJECTDOCSTRINGUNICODE",
        0x0006 => "PROJECTHELPFILEPATH",
        0x003D => "PROJECTHELPFILEPATH2",
        0x0007 => "PROJECTHELPCONTEXT",
        0x0008 => "PROJECTLIBFLAGS",
        0x0009 => "PROJECTVERSION",
        0x000C => "PROJECTCONSTANTS",
        0x003C => "PROJECTCONSTANTSUNICODE",
        0x0014 => "PROJECTLCIDINVOKE",
        // Some real-world files include this in the ProjectInformation list.
        0x004A => "PROJECTCOMPATVERSION",

        // ---- Reference records ----
        0x000D => "REFERENCEREGISTERED",
        0x000E => "REFERENCEPROJECT",
        0x0016 => "REFERENCENAME",
        0x002F => "REFERENCECONTROL",
        0x0030 => "REFERENCEEXTENDED",
        0x0033 => "REFERENCEORIGINAL",

        // ---- Module records ----
        0x000F => "PROJECTMODULES (ModuleCount)",
        0x0013 => "PROJECTCOOKIE",
        0x0010 => "PROJECTTERMINATOR (dir stream end)",

        0x0019 => "MODULENAME",
        0x0047 => "MODULENAMEUNICODE",

        0x001A => "MODULESTREAMNAME",
        0x0032 => "MODULESTREAMNAMEUNICODE",

        0x001C => "MODULEDOCSTRING",
        0x0048 => "MODULEDOCSTRINGUNICODE",

        0x001D => "MODULEHELPFILEPATH",
        0x001E => "MODULEHELPCONTEXT",

        0x0021 => "MODULETYPE (procedural TypeRecord.Id=0x0021)",
        0x0022 => "MODULETYPE (non-procedural TypeRecord.Id=0x0022)",
        0x0025 => "MODULEREADONLY",
        0x0028 => "MODULEPRIVATE",
        0x002B => "MODULETERMINATOR",
        0x002C => "MODULECOOKIE",
        0x0031 => "MODULETEXTOFFSET",

        // ---- Reserved marker Unicode record IDs (still seen) ----
        0x003E => "REFERENCENAMEUNICODE (reserved marker)",

        _ => return None,
    })
}
