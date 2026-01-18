use tauri::http::header::{HeaderName, HeaderValue, CONTENT_TYPE};
use tauri::http::{Response, StatusCode};

pub(crate) fn response(status: StatusCode, body: Vec<u8>) -> Response<Vec<u8>> {
    let mut response = Response::new(body);
    *response.status_mut() = status;
    response
}

pub(crate) fn response_with_content_type(
    status: StatusCode,
    content_type: &str,
    body: Vec<u8>,
) -> Response<Vec<u8>> {
    let mut response = response(status, body);
    insert_header_name_value(
        &mut response,
        CONTENT_TYPE,
        HeaderValue::from_str(content_type).ok(),
    );
    response
}

pub(crate) fn response_with_content_type_and_csp(
    status: StatusCode,
    content_type: &str,
    body: Vec<u8>,
    csp: Option<&str>,
) -> Response<Vec<u8>> {
    let mut response = response_with_content_type(status, content_type, body);
    if let Some(csp) = csp {
        insert_header(&mut response, "Content-Security-Policy", csp);
    }
    response
}

pub(crate) fn insert_header(response: &mut Response<Vec<u8>>, name: &str, value: &str) {
    let name = HeaderName::from_bytes(name.as_bytes()).ok();
    let value = HeaderValue::from_str(value).ok();
    insert_header_name_value(response, name, value);
}

fn insert_header_name_value(
    response: &mut Response<Vec<u8>>,
    name: impl Into<Option<HeaderName>>,
    value: Option<HeaderValue>,
) {
    let (Some(name), Some(value)) = (name.into(), value) else {
        return;
    };
    response.headers_mut().insert(name, value);
}

