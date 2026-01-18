#[derive(Debug, Clone)]
pub(crate) enum PatternToken {
    LiteralSeq(Vec<char>),
    AnyOne,
    AnyMany,
}

pub(crate) fn parse_search_pattern_folded(pattern: &str) -> Vec<PatternToken> {
    fn push_folded_literal(out: &mut Vec<char>, ch: char) {
        if ch.is_ascii() {
            out.push(ch.to_ascii_uppercase());
        } else {
            out.extend(ch.to_uppercase());
        }
    }

    let mut tokens = Vec::new();
    let mut literal: Vec<char> = Vec::new();

    let mut it = pattern.chars();
    while let Some(ch) = it.next() {
        if ch == '~' {
            match it.next() {
                Some(next) => push_folded_literal(&mut literal, next),
                None => literal.push('~'),
            }
            continue;
        }

        match ch {
            '*' => {
                if !literal.is_empty() {
                    tokens.push(PatternToken::LiteralSeq(std::mem::take(&mut literal)));
                }
                if !matches!(tokens.last(), Some(PatternToken::AnyMany)) {
                    tokens.push(PatternToken::AnyMany);
                }
            }
            '?' => {
                if !literal.is_empty() {
                    tokens.push(PatternToken::LiteralSeq(std::mem::take(&mut literal)));
                }
                tokens.push(PatternToken::AnyOne);
            }
            _ => push_folded_literal(&mut literal, ch),
        }
    }

    if !literal.is_empty() {
        tokens.push(PatternToken::LiteralSeq(literal));
    }

    tokens
}

pub(crate) fn min_required_hay_len(tokens: &[PatternToken]) -> usize {
    tokens
        .iter()
        .map(|t| match t {
            PatternToken::LiteralSeq(seq) => seq.len(),
            PatternToken::AnyOne => 1,
            PatternToken::AnyMany => 0,
        })
        .sum()
}

pub(crate) fn matches_pattern_with_memo(
    tokens: &[PatternToken],
    hay: &[char],
    start: usize,
    stride: usize,
    memo: &mut [Option<bool>],
) -> bool {
    debug_assert_eq!(stride, hay.len() + 1);
    debug_assert_eq!(memo.len(), (tokens.len() + 1) * stride);
    match_iter(tokens, hay, start, stride, memo)
}

fn match_iter(
    tokens: &[PatternToken],
    hay: &[char],
    start: usize,
    stride: usize,
    memo: &mut [Option<bool>],
) -> bool {
    #[derive(Clone, Copy)]
    struct Frame {
        tok_idx: usize,
        hay_idx: usize,
        expanded: bool,
    }

    let mut stack: Vec<Frame> = Vec::new();
    stack.push(Frame {
        tok_idx: 0,
        hay_idx: start,
        expanded: false,
    });

    while let Some(frame) = stack.pop() {
        debug_assert!(frame.hay_idx <= hay.len());
        debug_assert!(frame.tok_idx <= tokens.len());
        let memo_idx = frame.tok_idx * stride + frame.hay_idx;
        if memo[memo_idx].is_some() {
            continue;
        }

        if !frame.expanded {
            // Post-order evaluation: expand dependencies first, then compute this node.
            stack.push(Frame {
                expanded: true,
                ..frame
            });

            if frame.tok_idx == tokens.len() {
                continue;
            }
            match &tokens[frame.tok_idx] {
                PatternToken::LiteralSeq(seq) => {
                    let next_hay = frame.hay_idx + seq.len();
                    if next_hay <= hay.len() {
                        stack.push(Frame {
                            tok_idx: frame.tok_idx + 1,
                            hay_idx: next_hay,
                            expanded: false,
                        });
                    }
                }
                PatternToken::AnyOne => {
                    if frame.hay_idx < hay.len() {
                        stack.push(Frame {
                            tok_idx: frame.tok_idx + 1,
                            hay_idx: frame.hay_idx + 1,
                            expanded: false,
                        });
                    }
                }
                PatternToken::AnyMany => {
                    // '*' can match empty or one-or-more chars.
                    stack.push(Frame {
                        tok_idx: frame.tok_idx + 1,
                        hay_idx: frame.hay_idx,
                        expanded: false,
                    });
                    if frame.hay_idx < hay.len() {
                        stack.push(Frame {
                            tok_idx: frame.tok_idx,
                            hay_idx: frame.hay_idx + 1,
                            expanded: false,
                        });
                    }
                }
            }
            continue;
        }

        let result = if frame.tok_idx == tokens.len() {
            true
        } else {
            match &tokens[frame.tok_idx] {
                PatternToken::LiteralSeq(seq) => {
                    let next_hay = frame.hay_idx + seq.len();
                    if next_hay > hay.len() {
                        false
                    } else if hay[frame.hay_idx..next_hay] == *seq {
                        let child_idx = (frame.tok_idx + 1) * stride + next_hay;
                        memo[child_idx].unwrap_or(false)
                    } else {
                        false
                    }
                }
                PatternToken::AnyOne => {
                    if frame.hay_idx >= hay.len() {
                        false
                    } else {
                        let child_idx = (frame.tok_idx + 1) * stride + (frame.hay_idx + 1);
                        memo[child_idx].unwrap_or(false)
                    }
                }
                PatternToken::AnyMany => {
                    let empty_idx = (frame.tok_idx + 1) * stride + frame.hay_idx;
                    if memo[empty_idx].unwrap_or(false) {
                        true
                    } else if frame.hay_idx < hay.len() {
                        let more_idx = frame.tok_idx * stride + (frame.hay_idx + 1);
                        memo[more_idx].unwrap_or(false)
                    } else {
                        false
                    }
                }
            }
        };

        memo[memo_idx] = Some(result);
    }

    memo[0 * stride + start].unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::{
        matches_pattern_with_memo, min_required_hay_len, parse_search_pattern_folded, PatternToken,
    };

    fn alloc_memo(tokens_len: usize, stride: usize) -> Vec<Option<bool>> {
        let len = tokens_len
            .checked_add(1)
            .and_then(|n| n.checked_mul(stride))
            .unwrap_or_else(|| panic!("memo len overflow (tokens_len={tokens_len}, stride={stride})"));
        let mut memo: Vec<Option<bool>> = Vec::new();
        if memo.try_reserve_exact(len).is_err() {
            panic!("allocation failed (search_pattern memo, len={len})");
        }
        memo.resize(len, None);
        memo
    }

    fn fold_to_upper_chars(s: &str) -> Vec<char> {
        let mut out = Vec::new();
        for ch in s.chars() {
            if ch.is_ascii() {
                out.push(ch.to_ascii_uppercase());
            } else {
                out.extend(ch.to_uppercase());
            }
        }
        out
    }

    fn matches_at_start(pattern: &str, hay: &str) -> bool {
        let tokens = parse_search_pattern_folded(pattern);
        let hay = fold_to_upper_chars(hay);
        let stride = hay.len() + 1;
        let mut memo = alloc_memo(tokens.len(), stride);
        matches_pattern_with_memo(&tokens, &hay, 0, stride, &mut memo)
    }

    #[test]
    fn parse_collapses_consecutive_stars() {
        let tokens = parse_search_pattern_folded("a**b***c");
        assert!(matches!(
            tokens.as_slice(),
            [
                PatternToken::LiteralSeq(_),
                PatternToken::AnyMany,
                PatternToken::LiteralSeq(_),
                PatternToken::AnyMany,
                PatternToken::LiteralSeq(_),
            ]
        ));
    }

    #[test]
    fn parse_respects_excel_escape_with_tilde() {
        assert!(matches_at_start("~*~?~~", "*?~"));
        assert!(!matches_at_start("~*~?~~", "XX"));
    }

    #[test]
    fn min_required_hay_len_counts_literals_and_single_wildcards() {
        let tokens = parse_search_pattern_folded("*A?B*");
        assert_eq!(min_required_hay_len(&tokens), 3);
    }

    #[test]
    fn matcher_handles_basic_wildcards() {
        assert!(matches_at_start("A?C", "ABC"));
        assert!(matches_at_start("A?C", "AXC"));
        assert!(!matches_at_start("A?C", "AC"));

        assert!(matches_at_start("A*C", "AC"));
        assert!(matches_at_start("A*C", "ABBBBBBC"));
        assert!(!matches_at_start("A*C", "ABBBBBB"));
    }

    #[test]
    fn matcher_is_stack_safe_for_long_inputs() {
        let tokens = parse_search_pattern_folded("*");
        let mut hay: Vec<char> = Vec::new();
        if hay.try_reserve_exact(100_000).is_err() {
            panic!("allocation failed (search_pattern long hay)");
        }
        hay.resize(100_000, 'A');
        let stride = hay.len() + 1;
        let mut memo = alloc_memo(tokens.len(), stride);
        assert!(matches_pattern_with_memo(&tokens, &hay, 0, stride, &mut memo));
    }
}

