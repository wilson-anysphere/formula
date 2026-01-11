# Outbound TLS hardening & certificate pinning

Formula enterprise integrations that deliver data to customer-managed endpoints (starting with **SIEM delivery**) enforce:

- **Minimum TLS version:** TLS 1.3
- **Optional certificate pinning:** a server certificate **SHA-256** fingerprint allowlist

Certificate pinning is configured per-organization via `org_settings`:

- `certificate_pinning_enabled` (boolean)
- `certificate_pins` (jsonb array of SHA-256 fingerprints)

## What is pinned?

Pins are compared against the **leaf certificate** presented by the server during the TLS handshake. The client computes:

```
SHA-256( DER-encoded certificate bytes )
```

and compares it to the configured `certificate_pins` (case-insensitive; colons are ignored).

## Computing a pin

### From a PEM certificate file

```bash
openssl x509 -in server.crt -noout -fingerprint -sha256
```

Example output:

```
SHA256 Fingerprint=AA:BB:CC:...
```

Copy the hex portion after `=`. Pins may be stored with or without colons.

### From a live endpoint

```bash
HOST=example.com
PORT=443
openssl s_client -connect "$HOST:$PORT" -servername "$HOST" -showcerts </dev/null 2>/dev/null \
  | openssl x509 -noout -fingerprint -sha256
```

## Operational guidance

- During certificate rotation, add **multiple** fingerprints to `certificate_pins` to allow a smooth transition.
- If `certificate_pinning_enabled = true` but `certificate_pins` is empty/invalid, outbound delivery is treated as a configuration error and fails.

