use std::path::Path;

#[test]
fn collect_fixture_paths_includes_xlsb() {
    let tmpdir = tempfile::tempdir().unwrap();
    let root = tmpdir.path();

    std::fs::write(root.join("a.xlsx"), b"").unwrap();
    std::fs::write(root.join("b.xlsm"), b"").unwrap();
    std::fs::write(root.join("c.xlsb"), b"").unwrap();
    std::fs::write(root.join("d.txt"), b"").unwrap();

    let nested = root.join("nested");
    std::fs::create_dir_all(&nested).unwrap();
    std::fs::write(nested.join("e.xlsb"), b"").unwrap();

    let paths = xlsx_diff::collect_fixture_paths(Path::new(root)).unwrap();
    let names: Vec<String> = paths
        .iter()
        .map(|p| p.file_name().unwrap().to_string_lossy().to_string())
        .collect();

    assert_eq!(names, vec!["a.xlsx", "b.xlsm", "c.xlsb", "e.xlsb"]);
}

