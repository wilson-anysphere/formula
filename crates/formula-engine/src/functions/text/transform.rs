/// EXACT(text1, text2)
pub fn exact(text1: &str, text2: &str) -> bool {
    text1 == text2
}

/// CLEAN(text)
///
/// Removes non-printable characters (ASCII control codes 0-31 and DEL).
pub fn clean(text: &str) -> String {
    text.chars()
        .filter(|c| {
            let code = *c as u32;
            !(code <= 31 || code == 127)
        })
        .collect()
}

/// PROPER(text)
pub fn proper(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut new_word = true;
    for c in text.chars() {
        if c.is_alphabetic() {
            if new_word {
                out.extend(c.to_uppercase());
            } else {
                out.extend(c.to_lowercase());
            }
            new_word = false;
        } else {
            out.push(c);
            new_word = true;
        }
    }
    out
}
