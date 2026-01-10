pub fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

pub fn resolve_target(source_part: &str, target: &str) -> String {
    if let Some(target) = target.strip_prefix('/') {
        return normalize(target);
    }

    let base_dir = source_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    normalize(&format!("{base_dir}/{target}"))
}

fn normalize(path: &str) -> String {
    let mut out: Vec<&str> = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            other => out.push(other),
        }
    }
    out.join("/")
}
