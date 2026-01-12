/// Human-readable `VBA/dir` record names for debugging.
///
/// Record IDs are from **MS-OVBA §2.3.4 (dir Stream)**.
///
/// Notes:
/// - Some `VBA/dir` encodings emit Unicode/alternate strings as standalone records whose `data` is
///   UTF-16LE bytes (sometimes with an internal `u32le` length prefix).
/// - Some Unicode/alternate record IDs are "canonical" per MS-OVBA, while others are observed
///   non-canonical variants. This table includes both, with a bias toward the IDs used by our
///   hashing implementation/tests:
///   - Project strings: `PROJECTDOCSTRINGUNICODE` (0x0040), `PROJECTCONSTANTSUNICODE` (0x003C),
///     `PROJECTHELPFILEPATH2` (0x003D)
///   - Module strings: `MODULENAMEUNICODE` (0x0047), `MODULESTREAMNAMEUNICODE` (0x0032),
///     `MODULEDOCSTRINGUNICODE` (0x0048), `MODULEHELPFILEPATHUNICODE` (0x0049)
/// - `0x004A` is `PROJECTCOMPATVERSION`.
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
        0x0049 => "MODULEHELPFILEPATHUNICODE",
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
