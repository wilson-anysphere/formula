use std::io::{Cursor, Write};

use formula_vba::{compress_container, content_normalized_data};

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

fn build_two_module_project(module_order: [&str; 2]) -> Vec<u8> {
    // Distinct module source so we can assert byte ordering in ContentNormalizedData directly.
    let module_a_code = b"'MODULE-A\r\nSub A()\r\nEnd Sub\r\n";
    let module_b_code = b"'MODULE-B\r\nSub B()\r\nEnd Sub\r\n";

    let module_a_container = compress_container(module_a_code);
    let module_b_container = compress_container(module_b_code);

    // Build a minimal, spec-compliant decompressed `VBA/dir` stream that lists the modules in the
    // desired order. The critical part for this test is the order of module records:
    // MODULENAME (0x0019) starts each module record group.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME (required by some producers; harmless here).
        push_record(&mut out, 0x0004, b"VBAProject");

        for name in module_order {
            push_record(&mut out, 0x0019, name.as_bytes()); // MODULENAME

            // MODULESTREAMNAME + reserved u16.
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(name.as_bytes());
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);

            // MODULETYPE (standard)
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            // MODULETEXTOFFSET: our module stream is just the compressed container.
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        }

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Minimal OLE layout: `VBA/dir` + the module streams. Create the module streams in alphabetical
    // order (A then B) to ensure the tested ordering comes from `VBA/dir`, not OLE insertion order.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/ModuleA").expect("module A stream");
        s.write_all(&module_a_container).expect("write module A");
    }
    {
        let mut s = ole.create_stream("VBA/ModuleB").expect("module B stream");
        s.write_all(&module_b_container).expect("write module B");
    }

    ole.into_inner().into_inner()
}

#[test]
fn content_normalized_data_uses_module_record_order_from_dir_stream() {
    // Deliberately non-alphabetical order: B then A.
    let vba_bin = build_two_module_project(["ModuleB", "ModuleA"]);
    let normalized = content_normalized_data(&vba_bin).expect("content normalized data");

    let module_a_code = b"'MODULE-A\r\nSub A()\r\nEnd Sub\r\n";
    let module_b_code = b"'MODULE-B\r\nSub B()\r\nEnd Sub\r\n";

    let pos_b = find_subslice(&normalized, module_b_code).expect("ModuleB code should be present");
    let pos_a = find_subslice(&normalized, module_a_code).expect("ModuleA code should be present");
    assert!(
        pos_b < pos_a,
        "expected ModuleB bytes to appear before ModuleA bytes in ContentNormalizedData"
    );

    // Swapping module order in the `dir` stream should swap the order in the normalized data too.
    let vba_bin_swapped = build_two_module_project(["ModuleA", "ModuleB"]);
    let normalized_swapped =
        content_normalized_data(&vba_bin_swapped).expect("content normalized data");

    assert_ne!(
        normalized, normalized_swapped,
        "changing module stored order should change ContentNormalizedData"
    );

    let pos_a2 =
        find_subslice(&normalized_swapped, module_a_code).expect("ModuleA code should be present");
    let pos_b2 =
        find_subslice(&normalized_swapped, module_b_code).expect("ModuleB code should be present");
    assert!(
        pos_a2 < pos_b2,
        "expected ModuleA bytes to appear before ModuleB bytes when dir order is A then B"
    );
}

