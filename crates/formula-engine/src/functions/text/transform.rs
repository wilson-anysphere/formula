/// EXACT(text1, text2)
pub fn exact(text1: &str, text2: &str) -> bool {
    text1 == text2
}

/// CLEAN(text)
///
/// Removes non-printable characters (ASCII control codes 0-31 and DEL).
pub fn clean(text: &str) -> Result<String, crate::value::ErrorKind> {
    let mut out = String::new();
    if out.try_reserve_exact(text.len()).is_err() {
        debug_assert!(false, "allocation failed (clean, len={})", text.len());
        return Err(crate::value::ErrorKind::Num);
    }
    for c in text.chars() {
        let code = c as u32;
        if !(code <= 31 || code == 127) {
            out.push(c);
        }
    }
    Ok(out)
}

/// PROPER(text)
pub fn proper(text: &str) -> Result<String, crate::value::ErrorKind> {
    if text.is_ascii() {
        return proper_ascii(text);
    }

    let mut out = String::new();
    if out.try_reserve_exact(text.len()).is_err() {
        debug_assert!(false, "allocation failed (proper, len={})", text.len());
        return Err(crate::value::ErrorKind::Num);
    }
    let mut new_word = true;
    for c in text.chars() {
        if c.is_alphabetic() {
            if new_word {
                if c.is_ascii() {
                    out.push(c.to_ascii_uppercase());
                } else {
                    out.extend(c.to_uppercase());
                }
            } else {
                if c.is_ascii() {
                    out.push(c.to_ascii_lowercase());
                } else {
                    out.extend(c.to_lowercase());
                }
            }
            new_word = false;
        } else {
            out.push(c);
            new_word = true;
        }
    }
    Ok(out)
}

fn proper_ascii(text: &str) -> Result<String, crate::value::ErrorKind> {
    let mut out = String::new();
    if out.try_reserve_exact(text.len()).is_err() {
        debug_assert!(false, "allocation failed (proper_ascii, len={})", text.len());
        return Err(crate::value::ErrorKind::Num);
    }
    let mut new_word = true;
    for &b in text.as_bytes() {
        if b.is_ascii_alphabetic() {
            let c = b as char;
            out.push(if new_word {
                c.to_ascii_uppercase()
            } else {
                c.to_ascii_lowercase()
            });
            new_word = false;
        } else {
            out.push(b as char);
            new_word = true;
        }
    }
    Ok(out)
}
