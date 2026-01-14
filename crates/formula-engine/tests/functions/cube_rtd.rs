use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use formula_engine::{Engine, ErrorKind, ExternalDataProvider, Value};

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct RtdKey {
    prog_id: String,
    server: String,
    topics: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CubeValueKey {
    connection: String,
    tuples: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CubeMemberKey {
    connection: String,
    member_expression: String,
    caption: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CubeMemberPropertyKey {
    connection: String,
    member_expression_or_handle: String,
    property: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CubeRankedMemberKey {
    connection: String,
    set_expression_or_handle: String,
    rank: i64,
    caption: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CubeSetKey {
    connection: String,
    set_expression: String,
    caption: Option<String>,
    sort_order: Option<i64>,
    sort_by: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct CubeKpiMemberKey {
    connection: String,
    kpi_name: String,
    kpi_property: String,
    caption: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum ProviderCall {
    Rtd(RtdKey),
    CubeValue(CubeValueKey),
    CubeMember(CubeMemberKey),
    CubeMemberProperty(CubeMemberPropertyKey),
    CubeRankedMember(CubeRankedMemberKey),
    CubeSet(CubeSetKey),
    CubeSetCount(String),
    CubeKpiMember(CubeKpiMemberKey),
}

#[derive(Default)]
struct ProviderState {
    rtd: HashMap<RtdKey, Value>,
    cube_value: HashMap<CubeValueKey, Value>,
    cube_member: HashMap<CubeMemberKey, Value>,
    cube_member_property: HashMap<CubeMemberPropertyKey, Value>,
    cube_ranked_member: HashMap<CubeRankedMemberKey, Value>,
    cube_set: HashMap<CubeSetKey, Value>,
    cube_set_count: HashMap<String, Value>,
    cube_kpi_member: HashMap<CubeKpiMemberKey, Value>,
    calls: Vec<ProviderCall>,
}

#[derive(Default)]
struct TestExternalDataProvider {
    state: Mutex<ProviderState>,
}

impl TestExternalDataProvider {
    fn take_calls(&self) -> Vec<ProviderCall> {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        std::mem::take(&mut state.calls)
    }

    fn set_rtd(&self, prog_id: &str, server: &str, topics: &[&str], value: Value) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.rtd.insert(
            RtdKey {
                prog_id: prog_id.to_string(),
                server: server.to_string(),
                topics: topics.iter().map(|t| (*t).to_string()).collect(),
            },
            value,
        );
    }

    fn set_cube_value(&self, connection: &str, tuples: &[&str], value: Value) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.cube_value.insert(
            CubeValueKey {
                connection: connection.to_string(),
                tuples: tuples.iter().map(|t| (*t).to_string()).collect(),
            },
            value,
        );
    }

    fn set_cube_member(
        &self,
        connection: &str,
        member_expression: &str,
        caption: Option<&str>,
        value: Value,
    ) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.cube_member.insert(
            CubeMemberKey {
                connection: connection.to_string(),
                member_expression: member_expression.to_string(),
                caption: caption.map(|c| c.to_string()),
            },
            value,
        );
    }

    fn set_cube_member_property(
        &self,
        connection: &str,
        member_expression_or_handle: &str,
        property: &str,
        value: Value,
    ) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.cube_member_property.insert(
            CubeMemberPropertyKey {
                connection: connection.to_string(),
                member_expression_or_handle: member_expression_or_handle.to_string(),
                property: property.to_string(),
            },
            value,
        );
    }

    fn set_cube_ranked_member(
        &self,
        connection: &str,
        set_expression_or_handle: &str,
        rank: i64,
        caption: Option<&str>,
        value: Value,
    ) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.cube_ranked_member.insert(
            CubeRankedMemberKey {
                connection: connection.to_string(),
                set_expression_or_handle: set_expression_or_handle.to_string(),
                rank,
                caption: caption.map(|c| c.to_string()),
            },
            value,
        );
    }

    fn set_cube_set(
        &self,
        connection: &str,
        set_expression: &str,
        caption: Option<&str>,
        sort_order: Option<i64>,
        sort_by: Option<&str>,
        value: Value,
    ) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.cube_set.insert(
            CubeSetKey {
                connection: connection.to_string(),
                set_expression: set_expression.to_string(),
                caption: caption.map(|c| c.to_string()),
                sort_order,
                sort_by: sort_by.map(|s| s.to_string()),
            },
            value,
        );
    }

    fn set_cube_set_count(&self, set_expression_or_handle: &str, value: Value) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state
            .cube_set_count
            .insert(set_expression_or_handle.to_string(), value);
    }

    fn set_cube_kpi_member(
        &self,
        connection: &str,
        kpi_name: &str,
        kpi_property: &str,
        caption: Option<&str>,
        value: Value,
    ) {
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.cube_kpi_member.insert(
            CubeKpiMemberKey {
                connection: connection.to_string(),
                kpi_name: kpi_name.to_string(),
                kpi_property: kpi_property.to_string(),
                caption: caption.map(|c| c.to_string()),
            },
            value,
        );
    }
}

impl ExternalDataProvider for TestExternalDataProvider {
    fn rtd(&self, prog_id: &str, server: &str, topics: &[String]) -> Value {
        let key = RtdKey {
            prog_id: prog_id.to_string(),
            server: server.to_string(),
            topics: topics.to_vec(),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.calls.push(ProviderCall::Rtd(key.clone()));
        state
            .rtd
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_value(&self, connection: &str, tuples: &[String]) -> Value {
        let key = CubeValueKey {
            connection: connection.to_string(),
            tuples: tuples.to_vec(),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.calls.push(ProviderCall::CubeValue(key.clone()));
        state
            .cube_value
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_member(
        &self,
        connection: &str,
        member_expression: &str,
        caption: Option<&str>,
    ) -> Value {
        let key = CubeMemberKey {
            connection: connection.to_string(),
            member_expression: member_expression.to_string(),
            caption: caption.map(|c| c.to_string()),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.calls.push(ProviderCall::CubeMember(key.clone()));
        state
            .cube_member
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_member_property(
        &self,
        connection: &str,
        member_expression_or_handle: &str,
        property: &str,
    ) -> Value {
        let key = CubeMemberPropertyKey {
            connection: connection.to_string(),
            member_expression_or_handle: member_expression_or_handle.to_string(),
            property: property.to_string(),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state
            .calls
            .push(ProviderCall::CubeMemberProperty(key.clone()));
        state
            .cube_member_property
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_ranked_member(
        &self,
        connection: &str,
        set_expression_or_handle: &str,
        rank: i64,
        caption: Option<&str>,
    ) -> Value {
        let key = CubeRankedMemberKey {
            connection: connection.to_string(),
            set_expression_or_handle: set_expression_or_handle.to_string(),
            rank,
            caption: caption.map(|c| c.to_string()),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state
            .calls
            .push(ProviderCall::CubeRankedMember(key.clone()));
        state
            .cube_ranked_member
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_set(
        &self,
        connection: &str,
        set_expression: &str,
        caption: Option<&str>,
        sort_order: Option<i64>,
        sort_by: Option<&str>,
    ) -> Value {
        let key = CubeSetKey {
            connection: connection.to_string(),
            set_expression: set_expression.to_string(),
            caption: caption.map(|c| c.to_string()),
            sort_order,
            sort_by: sort_by.map(|s| s.to_string()),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.calls.push(ProviderCall::CubeSet(key.clone()));
        state
            .cube_set
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_set_count(&self, set_expression_or_handle: &str) -> Value {
        let key = set_expression_or_handle.to_string();
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.calls.push(ProviderCall::CubeSetCount(key.clone()));
        state
            .cube_set_count
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }

    fn cube_kpi_member(
        &self,
        connection: &str,
        kpi_name: &str,
        kpi_property: &str,
        caption: Option<&str>,
    ) -> Value {
        let key = CubeKpiMemberKey {
            connection: connection.to_string(),
            kpi_name: kpi_name.to_string(),
            kpi_property: kpi_property.to_string(),
            caption: caption.map(|c| c.to_string()),
        };
        let mut state = self.state.lock().expect("provider mutex poisoned");
        state.calls.push(ProviderCall::CubeKpiMember(key.clone()));
        state
            .cube_kpi_member
            .get(&key)
            .cloned()
            .unwrap_or(Value::Error(ErrorKind::NA))
    }
}

fn eval(engine: &mut Engine, formula: &str) -> Value {
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate();
    engine.get_cell_value("Sheet1", "A1")
}

#[test]
fn cube_and_rtd_return_na_without_provider() {
    let mut engine = Engine::new();
    engine.set_external_data_provider(None);

    let cases = [
        "=RTD(\"prog\",\"server\",\"topic\")",
        "=CUBEVALUE(\"conn\",\"tuple\")",
        "=CUBEMEMBER(\"conn\",\"member\")",
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=CUBERANKEDMEMBER(\"conn\",\"set\", 1)",
        "=CUBESET(\"conn\",\"set\")",
        "=CUBESETCOUNT(\"set\")",
        "=CUBEKPIMEMBER(\"conn\",\"kpi\",\"prop\")",
    ];

    for formula in cases {
        assert_eq!(eval(&mut engine, formula), Value::Error(ErrorKind::NA));
    }
}

#[test]
fn cube_and_rtd_delegate_to_provider() {
    let provider = Arc::new(TestExternalDataProvider::default());

    let mut engine = Engine::new();
    engine.set_external_data_provider(Some(provider.clone()));

    provider.set_rtd(
        "prog",
        "server",
        &["topic"],
        Value::Text("rtd-ok".to_string()),
    );
    assert_eq!(
        eval(&mut engine, "=RTD(\"prog\",\"server\",\"topic\")"),
        Value::Text("rtd-ok".to_string())
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::Rtd(RtdKey {
            prog_id: "prog".to_string(),
            server: "server".to_string(),
            topics: vec!["topic".to_string()],
        })]
    );

    provider.set_cube_value("conn", &["tuple1", "tuple2"], Value::Number(42.0));
    assert_eq!(
        eval(&mut engine, "=CUBEVALUE(\"conn\",\"tuple1\",\"tuple2\")"),
        Value::Number(42.0)
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeValue(CubeValueKey {
            connection: "conn".to_string(),
            tuples: vec!["tuple1".to_string(), "tuple2".to_string()],
        })]
    );

    provider.set_cube_member(
        "conn",
        "member_expr",
        None,
        Value::Text("member-handle".to_string()),
    );
    assert_eq!(
        eval(&mut engine, "=CUBEMEMBER(\"conn\",\"member_expr\")"),
        Value::Text("member-handle".to_string())
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeMember(CubeMemberKey {
            connection: "conn".to_string(),
            member_expression: "member_expr".to_string(),
            caption: None,
        })]
    );

    provider.set_cube_member_property(
        "conn",
        "member-handle",
        "prop",
        Value::Text("prop-value".to_string()),
    );
    assert_eq!(
        eval(
            &mut engine,
            "=CUBEMEMBERPROPERTY(\"conn\",\"member-handle\",\"prop\")"
        ),
        Value::Text("prop-value".to_string())
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeMemberProperty(CubeMemberPropertyKey {
            connection: "conn".to_string(),
            member_expression_or_handle: "member-handle".to_string(),
            property: "prop".to_string(),
        })]
    );

    provider.set_cube_ranked_member(
        "conn",
        "set-handle",
        7,
        Some("caption"),
        Value::Text("ranked-member".to_string()),
    );
    assert_eq!(
        eval(
            &mut engine,
            "=CUBERANKEDMEMBER(\"conn\",\"set-handle\",7,\"caption\")"
        ),
        Value::Text("ranked-member".to_string())
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeRankedMember(CubeRankedMemberKey {
            connection: "conn".to_string(),
            set_expression_or_handle: "set-handle".to_string(),
            rank: 7,
            caption: Some("caption".to_string()),
        })]
    );

    provider.set_cube_set(
        "conn",
        "set-expr",
        Some("caption"),
        Some(2),
        Some("sort-by"),
        Value::Text("set-handle".to_string()),
    );
    assert_eq!(
        eval(
            &mut engine,
            "=CUBESET(\"conn\",\"set-expr\",\"caption\",2,\"sort-by\")"
        ),
        Value::Text("set-handle".to_string())
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeSet(CubeSetKey {
            connection: "conn".to_string(),
            set_expression: "set-expr".to_string(),
            caption: Some("caption".to_string()),
            sort_order: Some(2),
            sort_by: Some("sort-by".to_string()),
        })]
    );

    provider.set_cube_set_count("set-handle", Value::Number(3.0));
    assert_eq!(
        eval(&mut engine, "=CUBESETCOUNT(\"set-handle\")"),
        Value::Number(3.0)
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeSetCount("set-handle".to_string())]
    );

    provider.set_cube_kpi_member(
        "conn",
        "kpi-name",
        "kpi-prop",
        Some("caption"),
        Value::Text("kpi-member".to_string()),
    );
    assert_eq!(
        eval(
            &mut engine,
            "=CUBEKPIMEMBER(\"conn\",\"kpi-name\",\"kpi-prop\",\"caption\")"
        ),
        Value::Text("kpi-member".to_string())
    );
    assert_eq!(
        provider.take_calls(),
        vec![ProviderCall::CubeKpiMember(CubeKpiMemberKey {
            connection: "conn".to_string(),
            kpi_name: "kpi-name".to_string(),
            kpi_property: "kpi-prop".to_string(),
            caption: Some("caption".to_string()),
        })]
    );
}

#[test]
fn cube_rtd_are_volatile_and_refresh_without_dirtying_cells() {
    let provider = Arc::new(TestExternalDataProvider::default());

    let mut engine = Engine::new();
    engine.set_external_data_provider(Some(provider.clone()));

    engine
        .set_cell_formula("Sheet1", "A1", "=RTD(\"prog\",\"server\",\"topic\")")
        .unwrap();

    provider.set_rtd("prog", "server", &["topic"], Value::Text("v1".to_string()));
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("v1".to_string())
    );

    provider.set_rtd("prog", "server", &["topic"], Value::Text("v2".to_string()));
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("v2".to_string())
    );
}

#[test]
fn getting_data_error_literal_parses_and_propagates() {
    let mut engine = Engine::new();

    assert_eq!(
        eval(&mut engine, "=#GETTING_DATA"),
        Value::Error(ErrorKind::GettingData)
    );

    assert_eq!(
        eval(&mut engine, "=1+#GETTING_DATA"),
        Value::Error(ErrorKind::GettingData)
    );
}

#[test]
fn provider_returned_getting_data_is_surfaced() {
    let provider = Arc::new(TestExternalDataProvider::default());

    let mut engine = Engine::new();
    engine.set_external_data_provider(Some(provider.clone()));

    provider.set_rtd(
        "prog",
        "server",
        &["topic"],
        Value::Error(ErrorKind::GettingData),
    );
    assert_eq!(
        eval(&mut engine, "=RTD(\"prog\",\"server\",\"topic\")"),
        Value::Error(ErrorKind::GettingData)
    );
}
