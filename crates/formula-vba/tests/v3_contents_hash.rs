use std::io::{Cursor, Write};

use formula_vba::{compress_container, v3_content_normalized_data};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

#[test]
fn v3_content_normalized_data_module_order_and_module_type_record_rules() {
    // We intentionally create:
    // - Two modules in non-alphabetical order in `VBA/dir` (ModuleB then ModuleA),
    // - Module streams created in alphabetical order in OLE (ModuleA then ModuleB),
    // to ensure ordering comes from `PROJECTMODULES.Modules` and not from sorting or OLE directory
    // enumeration order.
    //
    // Additionally, ModuleB uses a MODULETYPE record id 0x0021 (procedural) and ModuleA uses
    // 0x0022 (non-procedural). Only the 0x0021 record bytes (Id + Reserved) should be appended to
    // the V3 transcript.

    let module_proc_name = "ModuleB";
    let module_nonproc_name = "ModuleA";

    // Distinct module source markers so we can assert ordering in the output.
    // Do not end the module with a trailing newline so the V3 normalization logic doesn't append
    // an extra empty line before `HashModuleNameFlag` causes the module name to be appended.
    let module_proc_code = b"'PROC-MODULE\r\nSub Proc()\r\nEnd Sub";
    let module_nonproc_code = b"'NONPROC-MODULE\r\nSub NonProc()\r\nEnd Sub";

    let module_proc_container = compress_container(module_proc_code);
    let module_nonproc_container = compress_container(module_nonproc_code);

    // Build a minimal decompressed `VBA/dir` stream that lists modules in stored order.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0004, b"VBAProject"); // PROJECTNAME (unused by the algorithm here)

        // First module (procedural): ModuleB
        push_record(&mut out, 0x0019, module_proc_name.as_bytes()); // MODULENAME
        let mut proc_stream_name = Vec::new();
        proc_stream_name.extend_from_slice(module_proc_name.as_bytes());
        proc_stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &proc_stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (procedural; id=0x0021)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET

        // Second module (non-procedural): ModuleA
        push_record(&mut out, 0x0019, module_nonproc_name.as_bytes()); // MODULENAME
        let mut nonproc_stream_name = Vec::new();
        nonproc_stream_name.extend_from_slice(module_nonproc_name.as_bytes());
        nonproc_stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &nonproc_stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0022, &0u16.to_le_bytes()); // MODULETYPE (non-procedural; id=0x0022)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    // Assemble a minimal VBA project OLE file with `VBA/dir` and the module streams.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");
    ole.create_storage("VBA").expect("create VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    // Create module streams in alphabetical order (ModuleA then ModuleB) to ensure the tested
    // behavior comes from `VBA/dir`.
    {
        let mut s = ole
            .create_stream("VBA/ModuleA")
            .expect("non-procedural module stream");
        s.write_all(&module_nonproc_container)
            .expect("write non-procedural module");
    }
    {
        let mut s = ole
            .create_stream("VBA/ModuleB")
            .expect("procedural module stream");
        s.write_all(&module_proc_container)
            .expect("write procedural module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    // ---- Module ordering: stored order (ModuleB then ModuleA) ----
    let module_proc_expected = b"'PROC-MODULE\nSub Proc()\nEnd Sub\nModuleB\n";
    let module_nonproc_expected = b"'NONPROC-MODULE\nSub NonProc()\nEnd Sub\nModuleA\n";

    let pos_proc = find_subslice(&normalized, module_proc_expected)
        .expect("procedural module code present (normalized + module name)");
    let pos_nonproc = find_subslice(&normalized, module_nonproc_expected)
        .expect("non-procedural module code present (normalized + module name)");
    assert!(
        pos_proc < pos_nonproc,
        "expected ModuleB bytes to appear before ModuleA bytes in V3ContentNormalizedData"
    );

    // ---- MODULETYPE behavior: include TypeRecord bytes only for id 0x0021 ----
    let type_0021 = [0x21u8, 0x00, 0x00, 0x00]; // id=0x0021 (LE) + reserved(u16)=0x0000
    let type_0022 = [0x22u8, 0x00, 0x00, 0x00]; // id=0x0022 (LE) + reserved(u16)=0x0000

    let count_0021 = normalized
        .windows(type_0021.len())
        .filter(|w| *w == type_0021)
        .count();
    assert_eq!(
        count_0021, 1,
        "expected V3ContentNormalizedData to include TypeRecord (0x0021) exactly once"
    );
    assert!(
        !normalized
            .windows(type_0022.len())
            .any(|w| w == type_0022),
        "expected V3ContentNormalizedData to not include TypeRecord (0x0022) bytes"
    );

    // Ensure the included TypeRecord bytes are part of the first (procedural) module's contribution,
    // not appended after the second module.
    let pos_type_0021 =
        find_subslice(&normalized, &type_0021).expect("TypeRecord bytes should be present");
    assert!(
        pos_type_0021 < pos_nonproc,
        "expected TypeRecord bytes to appear before the non-procedural module contribution"
    );
}

#[test]
fn v3_content_normalized_data_module_source_normalization_defaultattributes_and_vb_name() {
    // MS-OVBA §2.4.2.5 module normalization edge cases:
    // - `Attribute VB_Name = ...` lines are skipped (case-insensitive prefix match).
    // - DefaultAttributes filtering is based on **byte-equality** against the 7 constant strings
    //   (i.e. NOT a case-insensitive compare).
    // - Output uses LF-only line endings.
    // - When at least one line is included, the module name is appended once, followed by LF
    //   (`HashModuleNameFlag` behavior).
    let module_name = "Module1";

    let module_code = concat!(
        // Skipped (VB_Name)
        "Attribute VB_Name = \"Module1\"\r\n",
        // Skipped (exact DefaultAttributes match)
        "Attribute VB_GlobalNameSpace = False\r\n",
        // Included: starts with `attribute` but differs by case from DefaultAttributes.
        "attribute VB_GlobalNameSpace = False\r\n",
        // Included: non-attribute code lines.
        "Option Explicit\r",
        "Sub Foo()\n",
        // No trailing newline; normalization should still include it with an LF.
        "End Sub",
    );

    let module_container = compress_container(module_code.as_bytes());

    // Minimal decompressed `VBA/dir` stream describing a single module at offset 0.
    let dir_decompressed = {
        let mut out = Vec::new();
        push_record(&mut out, 0x0019, module_name.as_bytes()); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(module_name.as_bytes());
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // Use a non-procedural MODULETYPE record id so there is no binary transcript prefix before
        // the normalized module lines.
        push_record(&mut out, 0x0022, &0u16.to_le_bytes());

        // MODULETEXTOFFSET: our module stream is just the compressed container.
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    let vba_bin = ole.into_inner().into_inner();

    let normalized = v3_content_normalized_data(&vba_bin).expect("V3ContentNormalizedData");

    let expected = concat!(
        "attribute VB_GlobalNameSpace = False\n",
        "Option Explicit\n",
        "Sub Foo()\n",
        "End Sub\n",
        "Module1\n",
    )
    .as_bytes()
    .to_vec();

    assert_eq!(normalized, expected);
    assert!(
        !normalized.contains(&b'\r'),
        "expected LF-only output (no CR bytes)"
    );
}
