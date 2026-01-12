/// Human-readable `VBA/dir` record names for debugging.
///
/// Record IDs are from **MS-OVBA §2.3.4 (dir Stream)**.
///
/// Notes (seen in the wild):
/// - Unicode string variants that matter for non-ASCII projects are most commonly:
///   - Project strings: `0x0040..=0x0043` (`PROJECT*UNICODE`)
///   - Module strings: `0x0047..=0x004A` (`MODULE*UNICODE`)
/// - Some producers / older layouts also emit "reserved marker" Unicode records such as `0x0032`
///   and `0x003C` (we label those too for convenience).
/// - `0x004A` is an ID collision: it can mean `PROJECTCOMPATVERSION` (project-level) or
///   `MODULEHELPFILEPATHUNICODE` (module-level), depending on where it appears.
pub fn record_name(id: u16) -> Option<&'static str> {
    Some(match id {
        // ---- Project information records ----
        0x0001 => "PROJECTSYSKIND",
        0x0002 => "PROJECTLCID",
        0x0003 => "PROJECTCODEPAGE",
        0x0004 => "PROJECTNAME",
        0x0040 => "PROJECTNAMEUNICODE",
        0x0005 => "PROJECTDOCSTRING",
        0x0041 => "PROJECTDOCSTRINGUNICODE",
        0x0006 => "PROJECTHELPFILEPATH",
        0x0042 => "PROJECTHELPFILEPATHUNICODE",
        0x0007 => "PROJECTHELPCONTEXT",
        0x0008 => "PROJECTLIBFLAGS",
        0x0009 => "PROJECTVERSION",
        0x000C => "PROJECTCONSTANTS",
        0x0043 => "PROJECTCONSTANTSUNICODE",
        0x0014 => "PROJECTLCIDINVOKE",
        // Some real-world files include this in the ProjectInformation list.
        0x004A => "PROJECTCOMPATVERSION / MODULEHELPFILEPATHUNICODE",

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
        0x0048 => "MODULESTREAMNAMEUNICODE",

        0x001B => "MODULEDOCSTRING",
        0x0049 => "MODULEDOCSTRINGUNICODE",

        0x001D => "MODULEHELPFILEPATH",
        0x001E => "MODULEHELPCONTEXT",

        0x0021 => "MODULETYPE (procedural TypeRecord.Id=0x0021)",
        0x0022 => "MODULETYPE (non-procedural TypeRecord.Id=0x0022)",
        0x0025 => "MODULEREADONLY",
        0x0028 => "MODULEPRIVATE",
        0x002B => "MODULETERMINATOR",
        0x002C => "MODULECOOKIE",
        0x0031 => "MODULETEXTOFFSET",

        // ---- Legacy / reserved marker Unicode record IDs (still seen) ----
        0x0032 => "MODULESTREAMNAMEUNICODE (reserved marker)",
        0x003C => "PROJECTCONSTANTSUNICODE (reserved marker)",
        0x003D => "PROJECTHELPFILEPATHUNICODE (reserved marker)",
        0x003E => "REFERENCENAMEUNICODE (reserved marker)",

        _ => return None,
    })
}

