use crate::ast::{
    BinOp, CallArg, CaseComparisonOp, CaseCondition, ConstDecl, Expr, LoopConditionKind, ParamDef,
    ProcedureDef, ProcedureKind, SelectCaseArm, Stmt, UnOp, VarDecl, VbaProgram, VbaType,
};
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

    fn peek_next_token(&self) -> Result<Token, VbaError> {
        let mut lexer = self.lexer.clone();
        lexer.next_token()
    }

    fn is_keyword(&self, kw: &str) -> bool {
        matches!(&self.lookahead.kind, TokenKind::Keyword(k) if k == kw)
    }

    fn is_keyword_seq(&self, first: &str, second: &str) -> Result<bool, VbaError> {
        if !self.is_keyword(first) {
            return Ok(false);
        }
        let next = self.peek_next_token()?;
        Ok(matches!(next.kind, TokenKind::Keyword(k) if k == second))
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

    fn eat_newlines(&mut self) -> Result<(), VbaError> {
        while matches!(self.lookahead.kind, TokenKind::Newline) {
            self.bump()?;
        }
        Ok(())
    }

    fn skip_to_end_of_line(&mut self) -> Result<(), VbaError> {
        while !matches!(self.lookahead.kind, TokenKind::Newline | TokenKind::Eof) {
            self.bump()?;
        }
        self.eat_newlines()?;
        Ok(())
    }

    fn parse_program(&mut self) -> Result<VbaProgram, VbaError> {
        let mut prog = VbaProgram::new();
        self.eat_newlines()?;

        while !matches!(self.lookahead.kind, TokenKind::Eof) {
            // Option/Attribute lines.
            if self.eat_keyword("option")? {
                // Only `Option Explicit` matters today.
                if self.eat_keyword("explicit")? {
                    prog.option_explicit = true;
                } else {
                    // `Option Compare`, `Option Base`, ...
                    while !matches!(self.lookahead.kind, TokenKind::Newline | TokenKind::Eof) {
                        self.bump()?;
                    }
                }
                self.eat_newlines()?;
                continue;
            }
            if self.eat_keyword("attribute")? {
                self.skip_to_end_of_line()?;
                continue;
            }

            // Ignore visibility modifiers (both on declarations and procedures).
            self.eat_keyword("private")?;
            self.eat_keyword("public")?;

            // Module-level declarations.
            if self.is_keyword("dim") {
                self.bump()?;
                let decls = self.parse_dim_decl_list()?;
                prog.module_vars.extend(decls);
                self.eat_newlines()?;
                continue;
            }
            if self.is_keyword("const") {
                self.bump()?;
                let decls = self.parse_const_decl_list()?;
                prog.module_consts.extend(decls);
                self.eat_newlines()?;
                continue;
            }
            // VBA also allows module-level variable declarations without `Dim`:
            //   `Public counter As Long`
            //   `Private ws As Worksheet`
            if matches!(self.lookahead.kind, TokenKind::Identifier(_)) {
                let decls = self.parse_dim_decl_list()?;
                prog.module_vars.extend(decls);
                self.eat_newlines()?;
                continue;
            }

            // Procedures.
            let kind = match &self.lookahead.kind {
                TokenKind::Keyword(k) if k == "sub" => {
                    self.bump()?;
                    ProcedureKind::Sub
                }
                TokenKind::Keyword(k) if k == "function" => {
                    self.bump()?;
                    ProcedureKind::Function
                }
                other => {
                    return Err(VbaError::Parse(format!(
                        "Expected `Sub` or `Function` but found {other:?} at {}:{}",
                        self.lookahead.line, self.lookahead.col
                    )))
                }
            };

            let name = self.parse_identifier()?;
            let params = self.parse_param_list()?;
            let return_type = if kind == ProcedureKind::Function && self.eat_keyword("as")? {
                Some(self.parse_type_name()?)
            } else {
                None
            };
            self.eat_newlines()?;

            let body = self.parse_block_until_end(kind)?;

            let def = ProcedureDef {
                name: name.clone(),
                kind,
                params,
                return_type,
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

    fn parse_type_name(&mut self) -> Result<VbaType, VbaError> {
        let name = match &self.lookahead.kind {
            TokenKind::Identifier(id) => {
                let id = id.clone();
                self.bump()?;
                id
            }
            TokenKind::Keyword(k) => {
                let k = k.clone();
                self.bump()?;
                k
            }
            other => {
                return Err(VbaError::Parse(format!(
                    "Expected type name but found {other:?} at {}:{}",
                    self.lookahead.line, self.lookahead.col
                )))
            }
        };

        if name.eq_ignore_ascii_case("integer") {
            return Ok(VbaType::Integer);
        }
        if name.eq_ignore_ascii_case("long") {
            return Ok(VbaType::Long);
        }
        if name.eq_ignore_ascii_case("double") {
            return Ok(VbaType::Double);
        }
        if name.eq_ignore_ascii_case("string") {
            return Ok(VbaType::String);
        }
        if name.eq_ignore_ascii_case("boolean") {
            return Ok(VbaType::Boolean);
        }
        if name.eq_ignore_ascii_case("date") {
            return Ok(VbaType::Date);
        }
        // Best-effort: treat unknown types as Variant. This keeps the interpreter permissive
        // for common declarations like `Dim ws As Worksheet` without needing a full type
        // system.
        Ok(VbaType::Variant)
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
            let by_ref = if self.eat_keyword("byval")? {
                false
            } else {
                self.eat_keyword("byref")?;
                true
            };

            let name = self.parse_identifier()?;
            let ty = if self.eat_keyword("as")? {
                Some(self.parse_type_name()?)
            } else {
                None
            };

            params.push(ParamDef { name, by_ref, ty });

            self.eat_newlines()?;
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                self.eat_newlines()?;
                continue;
            }
            break;
        }

        self.expect_token(TokenKind::RParen)?;
        Ok(params)
    }

    fn is_end_procedure(&self, kind: ProcedureKind) -> Result<bool, VbaError> {
        if !self.is_keyword("end") {
            return Ok(false);
        }
        let next = self.peek_next_token()?;
        Ok(match (kind, next.kind) {
            (ProcedureKind::Sub, TokenKind::Keyword(k)) if k == "sub" => true,
            (ProcedureKind::Function, TokenKind::Keyword(k)) if k == "function" => true,
            _ => false,
        })
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
            match k.as_str() {
                "call" => {
                    self.bump()?;
                    let expr = self.parse_expression()?;
                    return Ok(Stmt::ExprStmt(expr));
                }
                "debug" => return self.parse_debug_print(),
                "dim" => {
                    self.bump()?;
                    return Ok(Stmt::Dim(self.parse_dim_decl_list()?));
                }
                "const" => {
                    self.bump()?;
                    return Ok(Stmt::Const(self.parse_const_decl_list()?));
                }
                "set" => {
                    self.bump()?;
                    let target = self.parse_reference_expr()?;
                    self.expect_token(TokenKind::Eq)?;
                    let value = self.parse_expression()?;
                    return Ok(Stmt::Set { target, value });
                }
                "if" => return self.parse_if(),
                "for" => return self.parse_for(),
                "do" => return self.parse_do_loop(),
                "while" => return self.parse_while_wend(),
                "select" => return self.parse_select_case(),
                "with" => return self.parse_with(),
                "exit" => return self.parse_exit(),
                "on" => return self.parse_on_error(),
                "resume" => return self.parse_resume_stmt(),
                "goto" => {
                    self.bump()?;
                    let label = self.parse_identifier()?;
                    return Ok(Stmt::Goto(label));
                }
                _ => {}
            }
        }

        // Label: <Ident> :
        if let TokenKind::Identifier(label) = &self.lookahead.kind {
            let label = label.clone();
            self.bump()?; // consume ident
            if matches!(self.lookahead.kind, TokenKind::Colon) {
                self.bump()?;
                return Ok(Stmt::Label(label));
            }
            // Not a label: treat as expression start.
            let mut expr = Expr::Var(label);
            expr = self.parse_postfix(expr)?;
            return self.finish_stmt_from_expr(expr);
        }

        // Inside a `With` block it's common to start statements with `.Member`.
        // We need to parse the LHS without consuming `=` as an equality operator, so we use the
        // reference-expression parser and then decide whether this is an assignment or call.
        if matches!(self.lookahead.kind, TokenKind::Dot) {
            let expr = self.parse_reference_expr()?;
            return self.finish_stmt_from_expr(expr);
        }

        let expr = self.parse_expression()?;
        self.finish_stmt_from_expr(expr)
    }

    fn parse_exit(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("exit")?;
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
            TokenKind::Keyword(k) if k == "do" => {
                self.bump()?;
                Ok(Stmt::ExitDo)
            }
            other => Err(VbaError::Parse(format!(
                "Expected Sub/Function/For/Do after Exit but found {other:?} at {}:{}",
                self.lookahead.line, self.lookahead.col
            ))),
        }
    }

    fn parse_resume_stmt(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("resume")?;
        if self.eat_keyword("next")? {
            return Ok(Stmt::ResumeNext);
        }
        if matches!(self.lookahead.kind, TokenKind::Identifier(_)) {
            let label = self.parse_identifier()?;
            return Ok(Stmt::ResumeLabel(label));
        }
        Ok(Stmt::Resume)
    }

    fn finish_stmt_from_expr(&mut self, expr: Expr) -> Result<Stmt, VbaError> {
        if matches!(self.lookahead.kind, TokenKind::Eq) {
            self.bump()?;
            let value = self.parse_expression()?;
            return Ok(Stmt::Assign { target: expr, value });
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
        let expr = match &self.lookahead.kind {
            TokenKind::Identifier(name) => {
                let name = name.clone();
                self.bump()?;
                Expr::Var(name)
            }
            TokenKind::Dot => {
                // Inside `With`, assignments can use `.Member`.
                self.parse_primary()?
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

    fn parse_implicit_call_args(&mut self) -> Result<Vec<CallArg>, VbaError> {
        let mut args = Vec::new();
        loop {
            if matches!(
                self.lookahead.kind,
                TokenKind::Newline | TokenKind::Colon | TokenKind::Eof
            ) {
                break;
            }
            args.push(self.parse_call_arg(false)?);
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                continue;
            }
            break;
        }
        Ok(args)
    }

    fn parse_call_arg(&mut self, allow_missing: bool) -> Result<CallArg, VbaError> {
        if allow_missing && matches!(self.lookahead.kind, TokenKind::Comma | TokenKind::RParen) {
            return Ok(CallArg {
                name: None,
                expr: Expr::Missing,
            });
        }

        // Named argument: `Foo:=expr`
        let is_named = match &self.lookahead.kind {
            TokenKind::Identifier(_) | TokenKind::Keyword(_) => {
                matches!(self.peek_next_token()?.kind, TokenKind::ColonEq)
            }
            _ => false,
        };

        if is_named {
            let name = match &self.lookahead.kind {
                TokenKind::Identifier(id) => id.clone(),
                TokenKind::Keyword(k) => k.clone(),
                _ => unreachable!(),
            };
            self.bump()?;
            self.expect_token(TokenKind::ColonEq)?;
            let expr = self.parse_expression()?;
            return Ok(CallArg {
                name: Some(name),
                expr,
            });
        }

        Ok(CallArg {
            name: None,
            expr: self.parse_expression()?,
        })
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
            out.extend(self.parse_statement_list()?);
        }
        Ok(out)
    }

    fn parse_for(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("for")?;
        if self.eat_keyword("each")? {
            let var = self.parse_identifier()?;
            self.expect_keyword("in")?;
            let iterable = self.parse_expression()?;
            self.eat_newlines()?;
            let body = self.parse_block_until_keywords(&["next"])?;
            self.expect_keyword("next")?;
            if matches!(self.lookahead.kind, TokenKind::Identifier(_)) {
                self.bump()?;
            }
            return Ok(Stmt::ForEach {
                var,
                iterable,
                body,
            });
        }

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

    fn parse_do_loop(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("do")?;

        let pre_condition = if self.eat_keyword("while")? {
            Some((LoopConditionKind::While, self.parse_expression()?))
        } else if self.eat_keyword("until")? {
            Some((LoopConditionKind::Until, self.parse_expression()?))
        } else {
            None
        };

        self.eat_newlines()?;
        let body = self.parse_block_until_keywords(&["loop"])?;
        self.expect_keyword("loop")?;

        let post_condition = if self.eat_keyword("while")? {
            Some((LoopConditionKind::While, self.parse_expression()?))
        } else if self.eat_keyword("until")? {
            Some((LoopConditionKind::Until, self.parse_expression()?))
        } else {
            None
        };

        Ok(Stmt::DoLoop {
            pre_condition,
            post_condition,
            body,
        })
    }

    fn parse_while_wend(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("while")?;
        let cond = self.parse_expression()?;
        self.eat_newlines()?;
        let body = self.parse_block_until_keywords(&["wend"])?;
        self.expect_keyword("wend")?;
        Ok(Stmt::While { cond, body })
    }

    fn parse_select_case(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("select")?;
        self.expect_keyword("case")?;
        let expr = self.parse_expression()?;
        self.eat_newlines()?;

        let mut cases = Vec::new();
        let mut else_body = Vec::new();

        loop {
            self.eat_newlines()?;
            if self.is_keyword_seq("end", "select")? {
                break;
            }
            if !self.eat_keyword("case")? {
                return Err(VbaError::Parse(format!(
                    "Expected `Case` or `End Select` but found {:?} at {}:{}",
                    self.lookahead.kind, self.lookahead.line, self.lookahead.col
                )));
            }

            if self.eat_keyword("else")? {
                self.eat_newlines()?;
                else_body = self.parse_block_until_end_select()?;
                break;
            }

            let conditions = self.parse_case_conditions_list()?;
            self.eat_newlines()?;
            let body = self.parse_block_until_case_or_end_select()?;
            cases.push(SelectCaseArm { conditions, body });
        }

        self.expect_keyword("end")?;
        self.expect_keyword("select")?;

        Ok(Stmt::SelectCase {
            expr,
            cases,
            else_body,
        })
    }

    fn parse_block_until_case_or_end_select(&mut self) -> Result<Vec<Stmt>, VbaError> {
        let mut out = Vec::new();
        loop {
            self.eat_newlines()?;
            if self.is_keyword("case") || self.is_keyword_seq("end", "select")? {
                break;
            }
            if matches!(self.lookahead.kind, TokenKind::Eof) {
                return Err(VbaError::Parse("Unexpected EOF in Select Case".to_string()));
            }
            out.extend(self.parse_statement_list()?);
        }
        Ok(out)
    }

    fn parse_block_until_end_select(&mut self) -> Result<Vec<Stmt>, VbaError> {
        let mut out = Vec::new();
        loop {
            self.eat_newlines()?;
            if self.is_keyword_seq("end", "select")? {
                break;
            }
            if matches!(self.lookahead.kind, TokenKind::Eof) {
                return Err(VbaError::Parse("Unexpected EOF in Select Case".to_string()));
            }
            out.extend(self.parse_statement_list()?);
        }
        Ok(out)
    }

    fn parse_case_conditions_list(&mut self) -> Result<Vec<CaseCondition>, VbaError> {
        let mut conds = Vec::new();
        loop {
            conds.push(self.parse_case_condition()?);
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                continue;
            }
            break;
        }
        Ok(conds)
    }

    fn parse_case_condition(&mut self) -> Result<CaseCondition, VbaError> {
        if self.eat_keyword("is")? {
            let op = match &self.lookahead.kind {
                TokenKind::Eq => {
                    self.bump()?;
                    CaseComparisonOp::Eq
                }
                TokenKind::Ne => {
                    self.bump()?;
                    CaseComparisonOp::Ne
                }
                TokenKind::Lt => {
                    self.bump()?;
                    CaseComparisonOp::Lt
                }
                TokenKind::Le => {
                    self.bump()?;
                    CaseComparisonOp::Le
                }
                TokenKind::Gt => {
                    self.bump()?;
                    CaseComparisonOp::Gt
                }
                TokenKind::Ge => {
                    self.bump()?;
                    CaseComparisonOp::Ge
                }
                other => {
                    return Err(VbaError::Parse(format!(
                        "Expected comparison operator after `Is` but found {other:?} at {}:{}",
                        self.lookahead.line, self.lookahead.col
                    )))
                }
            };
            let rhs = self.parse_expression()?;
            return Ok(CaseCondition::Is { op, expr: rhs });
        }

        let start = self.parse_expression()?;
        if self.eat_keyword("to")? {
            let end = self.parse_expression()?;
            return Ok(CaseCondition::Range { start, end });
        }
        Ok(CaseCondition::Expr(start))
    }

    fn parse_with(&mut self) -> Result<Stmt, VbaError> {
        self.expect_keyword("with")?;
        let object = self.parse_expression()?;
        self.eat_newlines()?;
        let body = self.parse_block_until_end_with()?;
        self.expect_keyword("end")?;
        self.expect_keyword("with")?;
        Ok(Stmt::With { object, body })
    }

    fn parse_block_until_end_with(&mut self) -> Result<Vec<Stmt>, VbaError> {
        let mut out = Vec::new();
        loop {
            self.eat_newlines()?;
            if self.is_keyword_seq("end", "with")? {
                break;
            }
            if matches!(self.lookahead.kind, TokenKind::Eof) {
                return Err(VbaError::Parse("Unexpected EOF in With block".to_string()));
            }
            out.extend(self.parse_statement_list()?);
        }
        Ok(out)
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

    fn parse_dim_decl_list(&mut self) -> Result<Vec<VarDecl>, VbaError> {
        let mut vars = Vec::new();
        loop {
            let name = self.parse_identifier()?;
            let dims = if matches!(self.lookahead.kind, TokenKind::LParen) {
                self.parse_array_dims()?
            } else {
                Vec::new()
            };
            let ty = if self.eat_keyword("as")? {
                self.parse_type_name()?
            } else {
                VbaType::Variant
            };
            vars.push(VarDecl { name, ty, dims });
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                continue;
            }
            break;
        }
        Ok(vars)
    }

    fn parse_array_dims(&mut self) -> Result<Vec<crate::ast::ArrayDim>, VbaError> {
        self.expect_token(TokenKind::LParen)?;
        let mut dims = Vec::new();
        self.eat_newlines()?;
        if matches!(self.lookahead.kind, TokenKind::RParen) {
            self.bump()?;
            return Ok(dims);
        }
        loop {
            let first = self.parse_expression()?;
            let (lower, upper) = if self.eat_keyword("to")? {
                let upper = self.parse_expression()?;
                (Some(first), upper)
            } else {
                (None, first)
            };
            dims.push(crate::ast::ArrayDim { lower, upper });
            self.eat_newlines()?;
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                self.eat_newlines()?;
                continue;
            }
            break;
        }
        self.expect_token(TokenKind::RParen)?;
        Ok(dims)
    }

    fn parse_const_decl_list(&mut self) -> Result<Vec<ConstDecl>, VbaError> {
        let mut decls = Vec::new();
        loop {
            let name = self.parse_identifier()?;
            let ty = if self.eat_keyword("as")? {
                Some(self.parse_type_name()?)
            } else {
                None
            };
            self.expect_token(TokenKind::Eq)?;
            let value = self.parse_expression()?;
            decls.push(ConstDecl { name, ty, value });
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                continue;
            }
            break;
        }
        Ok(decls)
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
                TokenKind::Backslash => BinOp::IntDiv,
                TokenKind::Keyword(k) if k == "mod" => BinOp::Mod,
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
        self.parse_pow()
    }

    fn parse_pow(&mut self) -> Result<Expr, VbaError> {
        let mut expr = self.parse_primary()?;
        if matches!(self.lookahead.kind, TokenKind::Caret) {
            self.bump()?;
            let rhs = self.parse_pow()?; // right-associative
            expr = Expr::Binary {
                op: BinOp::Pow,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
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
                Expr::Call {
                    callee: Box::new(Expr::Var("__new".to_string())),
                    args: vec![CallArg {
                        name: None,
                        expr: Expr::Literal(VbaValue::String(class_name)),
                    }],
                }
            }
            TokenKind::Identifier(name) => {
                let name = name.clone();
                self.bump()?;
                Expr::Var(name)
            }
            TokenKind::Dot => {
                // With-default member access.
                self.bump()?;
                let member = match &self.lookahead.kind {
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
                            "Expected member name after `.` but found {other:?} at {}:{}",
                            self.lookahead.line, self.lookahead.col
                        )))
                    }
                };
                Expr::Member {
                    object: Box::new(Expr::With),
                    member,
                }
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

    fn parse_arg_list(&mut self) -> Result<Vec<CallArg>, VbaError> {
        self.expect_token(TokenKind::LParen)?;
        let mut args = Vec::new();
        self.eat_newlines()?;

        if matches!(self.lookahead.kind, TokenKind::RParen) {
            self.bump()?;
            return Ok(args);
        }

        loop {
            args.push(self.parse_call_arg(true)?);
            self.eat_newlines()?;
            if matches!(self.lookahead.kind, TokenKind::Comma) {
                self.bump()?;
                self.eat_newlines()?;
                // Allow trailing missing args like `Offset(, 1)`
                if matches!(self.lookahead.kind, TokenKind::RParen) {
                    args.push(CallArg {
                        name: None,
                        expr: Expr::Missing,
                    });
                    break;
                }
                continue;
            }
            break;
        }

        self.expect_token(TokenKind::RParen)?;
        Ok(args)
    }
}
