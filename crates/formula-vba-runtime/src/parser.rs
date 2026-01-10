use crate::ast::{BinOp, Expr, ParamDef, ProcedureDef, ProcedureKind, Stmt, UnOp, VbaProgram};
use crate::lexer::{Lexer, Token, TokenKind};
use crate::runtime::VbaError;
use crate::value::VbaValue;

pub fn parse_program(source: &str) -> Result<VbaProgram, VbaError> {
    let mut parser = Parser::new(source)?;
    parser.parse_program()
}

struct Parser<'a> {
    lexer: Lexer<'a>,
    lookahead: Token,
}

impl<'a> Parser<'a> {
    fn new(source: &'a str) -> Result<Self, VbaError> {
        let mut lexer = Lexer::new(source);
        let lookahead = lexer.next_token()?;
        Ok(Self { lexer, lookahead })
    }

    fn bump(&mut self) -> Result<Token, VbaError> {
        let current = self.lookahead.clone();
        self.lookahead = self.lexer.next_token()?;
        Ok(current)
    }

    fn expect_keyword(&mut self, kw: &str) -> Result<(), VbaError> {
        match &self.lookahead.kind {
            TokenKind::Keyword(k) if k == kw => {
                self.bump()?;
                Ok(())
            }
            other => Err(VbaError::Parse(format!(
                "Expected keyword `{kw}` but found {other:?} at {}:{}",
                self.lookahead.line, self.lookahead.col
            ))),
        }
    }

    fn eat_keyword(&mut self, kw: &str) -> Result<bool, VbaError> {
        match &self.lookahead.kind {
            TokenKind::Keyword(k) if k == kw => {
                self.bump()?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }

    fn eat_newlines(&mut self) -> Result<(), VbaError> {
        while matches!(self.lookahead.kind, TokenKind::Newline) {
            self.bump()?;
        }
        Ok(())
    }

    fn parse_program(&mut self) -> Result<VbaProgram, VbaError> {
        let mut prog = VbaProgram::new();
        self.eat_newlines()?;
        while !matches!(self.lookahead.kind, TokenKind::Eof) {
            // Skip Option/Attribute lines we don't handle.
            if self.eat_keyword("option")? {
                // Option Explicit etc.
                while !matches!(self.lookahead.kind, TokenKind::Newline | TokenKind::Eof) {
                    self.bump()?;
                }
                self.eat_newlines()?;
                continue;
            }

            // Ignore visibility modifiers.
            self.eat_keyword("private")?;
            self.eat_keyword("public")?;

            let kind = match &self.lookahead.kind {
                TokenKind::Keyword(k) if k == "sub" => {
                    self.bump()?;
                    ProcedureKind::Sub
                }
                TokenKind::Keyword(k) if k == "function" => {
                    self.bump()?;
                    ProcedureKind::Function
                }
                _ => {
                    return Err(VbaError::Parse(format!(
                        "Expected `Sub` or `Function` at {}:{}",
                        self.lookahead.line, self.lookahead.col
                    )))
                }
            };

            let name = self.parse_identifier()?;
            let params = self.parse_param_list()?;
            self.eat_newlines()?;

            let body = self.parse_block_until_end(kind)?;

            let def = ProcedureDef {
                name: name.clone(),
                kind,
                params,
                body,
            };
            prog.procedures.insert(name.to_ascii_lowercase(), def);
            self.eat_newlines()?;
        }
        Ok(prog)
    }

    fn parse_identifier(&mut self) -> Result<String, VbaError> {
        match &self.lookahead.kind {
            TokenKind::Identifier(name) => {
                let name = name.clone();
                self.bump()?;
                Ok(name)
            }
            other => Err(VbaError::Parse(format!(
                "Expected identifier but found {other:?} at {}:{}",
                self.lookahead.line, self.lookahead.col
            ))),
        }
    }

    fn parse_param_list(&mut self) -> Result<Vec<ParamDef>, VbaError> {
        if !matches!(self.lookahead.kind, TokenKind::LParen) {
            return Ok(Vec::new());
        }
        self.bump()?; // (
        let mut params = Vec::new();
        self.eat_newlines()?;
        if matches!(self.lookahead.kind, TokenKind::RParen) {
            self.bump()?;
            return Ok(params);
        }
        loop {
            // Optional ByVal/ByRef
            let by_ref = if self.eat_keyword("byval")? {
                false
            } else {
                self.eat_keyword("byref")?;
                true
            };
            let name = self.parse_identifier()?;
            // Optional `As Type`
            if self.eat_keyword("as")? {
                // Skip type name.
                match &self.lookahead.kind {
                    TokenKind::Identifier(_) | TokenKind::Keyword(_) => {
                        self.bump()?;
                    }
                    _ => {}
                }
            }
            params.push(ParamDef { name, by_ref });
            self.eat_newlines()?;
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                self.eat_newlines()?;
                continue;
            }
            break;
        }
        match &self.lookahead.kind {
            TokenKind::RParen => {
                self.bump()?;
                Ok(params)
            }
            other => Err(VbaError::Parse(format!(
                "Expected `)` but found {other:?} at {}:{}",
                self.lookahead.line, self.lookahead.col
            ))),
        }
    }

    fn parse_block_until_end(&mut self, kind: ProcedureKind) -> Result<Vec<Stmt>, VbaError> {
        let mut body = Vec::new();
        loop {
            self.eat_newlines()?;
            if matches!(self.lookahead.kind, TokenKind::Eof) {
                return Err(VbaError::Parse(
                    "Unexpected EOF in procedure body".to_string(),
                ));
            }
            if self.is_end_procedure(kind)? {
                break;
            }
            body.extend(self.parse_statement_list()?);
        }
        self.expect_keyword("end")?;
        match kind {
            ProcedureKind::Sub => self.expect_keyword("sub")?,
            ProcedureKind::Function => self.expect_keyword("function")?,
        }
        Ok(body)
    }

    fn is_end_procedure(&self, _kind: ProcedureKind) -> Result<bool, VbaError> {
        if let TokenKind::Keyword(k) = &self.lookahead.kind {
            if k == "end" {
                // peek next token without consuming? For simplicity, we assume `End Sub/Function`
                // appears correctly. We'll validate in parse_block_until_end.
                return Ok(true);
            }
        }
        Ok(false)
    }

    fn parse_statement_list(&mut self) -> Result<Vec<Stmt>, VbaError> {
        let mut stmts = Vec::new();
        while !matches!(self.lookahead.kind, TokenKind::Newline | TokenKind::Eof) {
            if matches!(self.lookahead.kind, TokenKind::Colon) {
                self.bump()?;
                continue;
            }
            stmts.push(self.parse_statement()?);
            if matches!(self.lookahead.kind, TokenKind::Colon) {
                self.bump()?;
                continue;
            }
            break;
        }
        self.eat_newlines()?;
        Ok(stmts)
    }

    fn parse_statement(&mut self) -> Result<Stmt, VbaError> {
        if let TokenKind::Keyword(k) = &self.lookahead.kind {
            if k == "call" {
                self.bump()?;
                let expr = self.parse_expression()?;
                return Ok(Stmt::ExprStmt(expr));
            }
            if k == "debug" {
                return self.parse_debug_print();
            }
        }
        // Label: <Ident> :
        if let TokenKind::Identifier(label) = &self.lookahead.kind {
            let label = label.clone();
            // we need to peek next token; simplest: clone parser state? We'll do cheap by lexing ahead.
            // We'll instead rely on `parse_statement_list` handling `:` after statement; so labels need
            // to be parsed when an identifier is followed by `:`.
            // We'll do: consume ident, if next is colon then return Label.
            self.bump()?; // consume ident
            if matches!(self.lookahead.kind, TokenKind::Colon) {
                self.bump()?; // consume colon
                return Ok(Stmt::Label(label));
            }
            // Not a label: restore by treating it as start of expression statement / assignment.
            // This is painful. We'll cheat by using a small re-lex: we can't easily un-bump.
            // Instead we keep `saved` and build expression starting with Var(saved).
            // We'll continue parsing with expr having already consumed identifier.
            // We'll parse a suffix chain (member/call/index) and then see if assignment follows.
            let mut expr = Expr::Var(label);
            expr = self.parse_postfix(expr)?;
            return self.finish_stmt_from_expr(expr);
        }

        match &self.lookahead.kind {
            TokenKind::Keyword(k) if k == "dim" => {
                self.bump()?;
                let mut vars = Vec::new();
                loop {
                    let name = match &self.lookahead.kind {
                        TokenKind::Identifier(n) => {
                            let n = n.clone();
                            self.bump()?;
                            n
                        }
                        other => {
                            return Err(VbaError::Parse(format!(
                                "Expected identifier in Dim but found {other:?} at {}:{}",
                                self.lookahead.line, self.lookahead.col
                            )))
                        }
                    };
                    // Optional array dims: (..)
                    if matches!(self.lookahead.kind, TokenKind::LParen) {
                        // Consume until ')'
                        self.bump()?;
                        while !matches!(self.lookahead.kind, TokenKind::RParen | TokenKind::Eof) {
                            self.bump()?;
                        }
                        if matches!(self.lookahead.kind, TokenKind::RParen) {
                            self.bump()?;
                        }
                    }
                    // Optional `As Type` - ignore
                    if self.eat_keyword("as")? {
                        if matches!(
                            self.lookahead.kind,
                            TokenKind::Identifier(_) | TokenKind::Keyword(_)
                        ) {
                            self.bump()?;
                        }
                    }
                    vars.push(name);
                    if matches!(self.lookahead.kind, TokenKind::Comma) {
                        self.bump()?;
                        continue;
                    }
                    break;
                }
                Ok(Stmt::Dim(vars))
            }
            TokenKind::Keyword(k) if k == "set" => {
                self.bump()?;
                let target = self.parse_reference_expr()?;
                self.expect_token(TokenKind::Eq)?;
                let value = self.parse_expression()?;
                Ok(Stmt::Set { target, value })
            }
            TokenKind::Keyword(k) if k == "if" => self.parse_if(),
            TokenKind::Keyword(k) if k == "for" => self.parse_for(),
            TokenKind::Keyword(k) if k == "do" => self.parse_do_while(),
            TokenKind::Keyword(k) if k == "exit" => {
                self.bump()?;
                match &self.lookahead.kind {
                    TokenKind::Keyword(k) if k == "sub" => {
                        self.bump()?;
                        Ok(Stmt::ExitSub)
                    }
                    TokenKind::Keyword(k) if k == "function" => {
                        self.bump()?;
                        Ok(Stmt::ExitFunction)
                    }
                    TokenKind::Keyword(k) if k == "for" => {
                        self.bump()?;
                        Ok(Stmt::ExitFor)
                    }
                    other => Err(VbaError::Parse(format!(
                        "Expected Sub/Function/For after Exit but found {other:?} at {}:{}",
                        self.lookahead.line, self.lookahead.col
                    ))),
                }
            }
            TokenKind::Keyword(k) if k == "on" => self.parse_on_error(),
            TokenKind::Keyword(k) if k == "goto" => {
                self.bump()?;
                let label = self.parse_identifier()?;
                Ok(Stmt::Goto(label))
            }
            _ => {
                let expr = self.parse_expression()?;
                self.finish_stmt_from_expr(expr)
            }
        }
    }

    fn finish_stmt_from_expr(&mut self, expr: Expr) -> Result<Stmt, VbaError> {
        // Assignment?
        if matches!(self.lookahead.kind, TokenKind::Eq) {
            self.bump()?;
            let value = self.parse_expression()?;
            return Ok(Stmt::Assign {
                target: expr,
                value,
            });
        }
        // Implicit call (no parentheses) - parse argument list until end-of-line/colon.
        if !matches!(
            self.lookahead.kind,
            TokenKind::Newline | TokenKind::Colon | TokenKind::Eof
        ) {
            let args = self.parse_implicit_call_args()?;
            return Ok(Stmt::ExprStmt(Expr::Call {
                callee: Box::new(expr),
                args,
            }));
        }
        Ok(Stmt::ExprStmt(expr))
    }

    fn parse_reference_expr(&mut self) -> Result<Expr, VbaError> {
        // Similar to `parse_primary`, but intentionally excludes binary operators so
        // `Set c = ...` doesn't parse `c = ...` as an equality expression.
        let expr = match &self.lookahead.kind {
            TokenKind::Identifier(name) => {
                let name = name.clone();
                self.bump()?;
                Expr::Var(name)
            }
            other => {
                return Err(VbaError::Parse(format!(
                    "Expected identifier but found {other:?} at {}:{}",
                    self.lookahead.line, self.lookahead.col
                )))
            }
        };
        self.parse_postfix(expr)
    }

    fn parse_implicit_call_args(&mut self) -> Result<Vec<Expr>, VbaError> {
        let mut args = Vec::new();
        loop {
            if matches!(
                self.lookahead.kind,
                TokenKind::Newline | TokenKind::Colon | TokenKind::Eof
            ) {
                break;
            }
            args.push(self.parse_expression()?);
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                continue;
            }
            break;
        }
        Ok(args)
    }

    fn parse_debug_print(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("debug")?;
        self.expect_token(TokenKind::Dot)?;
        match &self.lookahead.kind {
            TokenKind::Keyword(k) if k == "print" => {
                self.bump()?;
            }
            TokenKind::Identifier(id) if id.eq_ignore_ascii_case("print") => {
                self.bump()?;
            }
            other => {
                return Err(VbaError::Parse(format!(
                    "Expected `Print` after `Debug.` but found {other:?} at {}:{}",
                    self.lookahead.line, self.lookahead.col
                )))
            }
        }
        // `Debug.Print` accepts an optional expression list, separated by commas.
        let args = self.parse_implicit_call_args()?;
        Ok(Stmt::ExprStmt(Expr::Call {
            callee: Box::new(Expr::Var("DebugPrint".to_string())),
            args,
        }))
    }

    fn parse_if(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("if")?;
        let cond = self.parse_expression()?;
        self.expect_keyword("then")?;

        // Single-line If: If cond Then <stmt> [Else <stmt>]
        if !matches!(self.lookahead.kind, TokenKind::Newline) {
            let then_stmts = self.parse_inline_statement_list(&["else"])?;
            let mut else_body = Vec::new();
            if self.eat_keyword("else")? {
                else_body = self.parse_inline_statement_list(&[])?;
            }
            return Ok(Stmt::If {
                cond,
                then_body: then_stmts,
                elseifs: Vec::new(),
                else_body,
            });
        }

        self.eat_newlines()?;
        let then_body = self.parse_block_until_keywords(&["elseif", "else", "end"])?;

        let mut elseifs = Vec::new();
        while self.eat_keyword("elseif")? {
            let elseif_cond = self.parse_expression()?;
            self.expect_keyword("then")?;
            self.eat_newlines()?;
            let body = self.parse_block_until_keywords(&["elseif", "else", "end"])?;
            elseifs.push((elseif_cond, body));
        }

        let else_body = if self.eat_keyword("else")? {
            self.eat_newlines()?;
            self.parse_block_until_keywords(&["end"])?
        } else {
            Vec::new()
        };

        self.expect_keyword("end")?;
        self.expect_keyword("if")?;

        Ok(Stmt::If {
            cond,
            then_body,
            elseifs,
            else_body,
        })
    }

    fn parse_inline_statement_list(&mut self, stop_kws: &[&str]) -> Result<Vec<Stmt>, VbaError> {
        let mut stmts = Vec::new();
        while !matches!(self.lookahead.kind, TokenKind::Newline | TokenKind::Eof) {
            if matches!(self.lookahead.kind, TokenKind::Colon) {
                self.bump()?;
                continue;
            }
            if let TokenKind::Keyword(k) = &self.lookahead.kind {
                if stop_kws.iter().any(|kw| *kw == k) {
                    break;
                }
            }
            stmts.push(self.parse_statement()?);
            if matches!(self.lookahead.kind, TokenKind::Colon) {
                self.bump()?;
                continue;
            }
            break;
        }
        Ok(stmts)
    }

    fn parse_block_until_keywords(&mut self, kws: &[&str]) -> Result<Vec<Stmt>, VbaError> {
        let mut out = Vec::new();
        loop {
            self.eat_newlines()?;
            if let TokenKind::Keyword(k) = &self.lookahead.kind {
                if kws.iter().any(|kw| *kw == k) {
                    break;
                }
            }
            if matches!(self.lookahead.kind, TokenKind::Eof) {
                return Err(VbaError::Parse("Unexpected EOF in block".to_string()));
            }
            let stmts = self.parse_statement_list()?;
            out.extend(stmts);
        }
        Ok(out)
    }

    fn parse_for(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("for")?;
        let var = self.parse_identifier()?;
        self.expect_token(TokenKind::Eq)?;
        let start = self.parse_expression()?;
        self.expect_keyword("to")?;
        let end = self.parse_expression()?;
        let step = if self.eat_keyword("step")? {
            Some(self.parse_expression()?)
        } else {
            None
        };
        self.eat_newlines()?;
        let body = self.parse_block_until_keywords(&["next"])?;
        self.expect_keyword("next")?;
        // Optional loop variable
        if matches!(self.lookahead.kind, TokenKind::Identifier(_)) {
            self.bump()?;
        }
        Ok(Stmt::For {
            var,
            start,
            end,
            step,
            body,
        })
    }

    fn parse_do_while(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("do")?;
        self.expect_keyword("while")?;
        let cond = self.parse_expression()?;
        self.eat_newlines()?;
        let body = self.parse_block_until_keywords(&["loop"])?;
        self.expect_keyword("loop")?;
        Ok(Stmt::DoWhile { cond, body })
    }

    fn parse_on_error(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("on")?;
        self.expect_keyword("error")?;
        if self.eat_keyword("resume")? {
            self.expect_keyword("next")?;
            return Ok(Stmt::OnErrorResumeNext);
        }
        if self.eat_keyword("goto")? {
            match &self.lookahead.kind {
                TokenKind::Number(n) if (*n - 0.0).abs() < f64::EPSILON => {
                    self.bump()?;
                    return Ok(Stmt::OnErrorGoto0);
                }
                TokenKind::Identifier(_) => {
                    let label = self.parse_identifier()?;
                    return Ok(Stmt::OnErrorGotoLabel(label));
                }
                other => {
                    return Err(VbaError::Parse(format!(
                        "Expected 0 or label after `On Error GoTo` but found {other:?} at {}:{}",
                        self.lookahead.line, self.lookahead.col
                    )))
                }
            }
        }
        Err(VbaError::Parse(format!(
            "Unsupported `On Error` form at {}:{}",
            self.lookahead.line, self.lookahead.col
        )))
    }

    fn expect_token(&mut self, kind: TokenKind) -> Result<(), VbaError> {
        if std::mem::discriminant(&self.lookahead.kind) == std::mem::discriminant(&kind) {
            self.bump()?;
            Ok(())
        } else {
            Err(VbaError::Parse(format!(
                "Expected token {kind:?} but found {:?} at {}:{}",
                self.lookahead.kind, self.lookahead.line, self.lookahead.col
            )))
        }
    }

    // -------- Expressions (precedence climbing) --------
    fn parse_expression(&mut self) -> Result<Expr, VbaError> {
        self.parse_or()
    }

    fn parse_or(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_and()?;
        while self.eat_keyword("or")? {
            let rhs = self.parse_and()?;
            expr = Expr::Binary {
                op: BinOp::Or,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_and(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_cmp()?;
        while self.eat_keyword("and")? {
            let rhs = self.parse_cmp()?;
            expr = Expr::Binary {
                op: BinOp::And,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_cmp(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_concat()?;
        loop {
            let op = match &self.lookahead.kind {
                TokenKind::Eq => BinOp::Eq,
                TokenKind::Ne => BinOp::Ne,
                TokenKind::Lt => BinOp::Lt,
                TokenKind::Le => BinOp::Le,
                TokenKind::Gt => BinOp::Gt,
                TokenKind::Ge => BinOp::Ge,
                _ => break,
            };
            self.bump()?;
            let rhs = self.parse_concat()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_concat(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_add()?;
        while matches!(self.lookahead.kind, TokenKind::Amp) {
            self.bump()?;
            let rhs = self.parse_add()?;
            expr = Expr::Binary {
                op: BinOp::Concat,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_add(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_mul()?;
        loop {
            let op = match &self.lookahead.kind {
                TokenKind::Plus => BinOp::Add,
                TokenKind::Minus => BinOp::Sub,
                _ => break,
            };
            self.bump()?;
            let rhs = self.parse_mul()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_mul(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_unary()?;
        loop {
            let op = match &self.lookahead.kind {
                TokenKind::Star => BinOp::Mul,
                TokenKind::Slash => BinOp::Div,
                _ => break,
            };
            self.bump()?;
            let rhs = self.parse_unary()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_unary(&mut self) -> Result<Expr, VbaError> {
        if matches!(self.lookahead.kind, TokenKind::Minus) {
            self.bump()?;
            let expr = self.parse_unary()?;
            return Ok(Expr::Unary {
                op: UnOp::Neg,
                expr: Box::new(expr),
            });
        }
        if self.eat_keyword("not")? {
            let expr = self.parse_unary()?;
            return Ok(Expr::Unary {
                op: UnOp::Not,
                expr: Box::new(expr),
            });
        }
        self.parse_primary()
    }

    fn parse_primary(&mut self) -> Result<Expr, VbaError> {
        let expr = match &self.lookahead.kind {
            TokenKind::Number(n) => {
                let v = *n;
                self.bump()?;
                Expr::Literal(VbaValue::Double(v))
            }
            TokenKind::String(s) => {
                let s = s.clone();
                self.bump()?;
                Expr::Literal(VbaValue::String(s))
            }
            TokenKind::Keyword(k) if k == "true" => {
                self.bump()?;
                Expr::Literal(VbaValue::Boolean(true))
            }
            TokenKind::Keyword(k) if k == "false" => {
                self.bump()?;
                Expr::Literal(VbaValue::Boolean(false))
            }
            TokenKind::Keyword(k) if k == "nothing" => {
                self.bump()?;
                Expr::Literal(VbaValue::Null)
            }
            TokenKind::Keyword(k) if k == "new" => {
                self.bump()?;
                let class_name = match &self.lookahead.kind {
                    TokenKind::Identifier(name) => {
                        let name = name.clone();
                        self.bump()?;
                        name
                    }
                    TokenKind::Keyword(k) => {
                        let name = k.clone();
                        self.bump()?;
                        name
                    }
                    other => {
                        return Err(VbaError::Parse(format!(
                            "Expected class name after `New` but found {other:?} at {}:{}",
                            self.lookahead.line, self.lookahead.col
                        )))
                    }
                };
                // Model `New <Class>` as an internal constructor call. This keeps the AST small
                // while allowing the runtime to gate object creation through its sandbox.
                Expr::Call {
                    callee: Box::new(Expr::Var("__new".to_string())),
                    args: vec![Expr::Literal(VbaValue::String(class_name))],
                }
            }
            TokenKind::Identifier(name) => {
                let name = name.clone();
                self.bump()?;
                Expr::Var(name)
            }
            TokenKind::LParen => {
                self.bump()?;
                let expr = self.parse_expression()?;
                self.expect_token(TokenKind::RParen)?;
                expr
            }
            other => {
                return Err(VbaError::Parse(format!(
                    "Unexpected token {other:?} in expression at {}:{}",
                    self.lookahead.line, self.lookahead.col
                )))
            }
        };
        self.parse_postfix(expr)
    }

    fn parse_postfix(&mut self, mut expr: Expr) -> Result<Expr, VbaError> {
        loop {
            match &self.lookahead.kind {
                TokenKind::Dot => {
                    self.bump()?;
                    let member = match &self.lookahead.kind {
                        TokenKind::Identifier(name) => {
                            let name = name.clone();
                            self.bump()?;
                            name
                        }
                        TokenKind::Keyword(k) => {
                            // allow `.Value` etc as keyword-ish.
                            let name = k.clone();
                            self.bump()?;
                            name
                        }
                        other => {
                            return Err(VbaError::Parse(format!(
                                "Expected member name after `.` but found {other:?} at {}:{}",
                                self.lookahead.line, self.lookahead.col
                            )))
                        }
                    };
                    expr = Expr::Member {
                        object: Box::new(expr),
                        member,
                    };
                }
                TokenKind::LParen => {
                    // Call or index. For simplicity treat as call always; runtime may interpret
                    // `arr(i)` as Index if the callee is an array variable.
                    let args = self.parse_arg_list()?;
                    expr = Expr::Call {
                        callee: Box::new(expr),
                        args,
                    };
                }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_arg_list(&mut self) -> Result<Vec<Expr>, VbaError> {
        self.expect_token(TokenKind::LParen)?;
        let mut args = Vec::new();
        self.eat_newlines()?;
        if matches!(self.lookahead.kind, TokenKind::RParen) {
            self.bump()?;
            return Ok(args);
        }
        loop {
            args.push(self.parse_expression()?);
            self.eat_newlines()?;
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                self.eat_newlines()?;
                continue;
            }
            break;
        }
        self.expect_token(TokenKind::RParen)?;
        Ok(args)
    }
}
