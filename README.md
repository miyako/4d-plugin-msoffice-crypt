![version](https://img.shields.io/badge/version-17%2B-3E8B93)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-msoffice-crypt)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-msoffice-crypt/total)

# 4d-plugin-msoffice-crypt

`msoffice crypt` encrypts and decrypts Office Open XML documents (`.xlsx`, `.docx`, `.pptx`, and their template/macro-enabled variants) using the password-protection scheme defined by MS-OFFCRYPTO — the same CFB-container-wrapped encryption Microsoft Office itself uses. Both commands take the document as a `Blob` and a set of options as an `Object`, and return a status `Object` describing the outcome.

| Command | Returns | Purpose |
|---|---|---|
| [`msoffice encrypt`](#msoffice-encrypt) | Object | Encrypts an Office Open XML document, in place, with a password. |
| [`msoffice decrypt`](#msoffice-decrypt) | Object | Decrypts a password-protected Office Open XML document, in place. |

**Platforms:** Windows, macOS

---

## Requirements & platform notes

- **The document parameter is mutated in place.** Both commands write their result back into the same `Blob` variable you passed as parameter 1 — there is no separate "output" parameter or return value carrying the document. The status object only carries metadata (`success`, `warning`, `error`, etc.), never the document bytes themselves.
- **A password is required to do anything.** If none of `password`, `password_hex`, or `password_uni` is present in the options object (including if the whole `options` parameter is omitted), the command returns `{success: false}` and leaves the input `Blob` untouched — silently, with no `error` field set, unless the input itself couldn't be read at all (see [Error handling](#error-handling--troubleshooting)).
- **The three password forms are alternatives, not additive**, and are resolved in this priority order: `password` first, then `password_hex`, then `password_uni`. If more than one is present in the same options object, only the highest-priority one present is actually used. The same rule and order applies to the `secret`/`secret_hex`/`secret_uni` trio.
- **`mode` only affects `msoffice encrypt`.** Passing it to `msoffice decrypt` has no effect.
- **Warnings don't block the operation.** If you call `msoffice encrypt` on data that's already encrypted, or `msoffice decrypt` on data that isn't encrypted, the command still attempts the operation and reports a `warning` field rather than refusing outright — check `warning` if you want to guard against this.
- Both commands are declared thread-safe in the plugin manifest, so they can be called from any 4D process, including a worker/preemptive process.
- No platform-specific behavior differences are exposed to the 4D developer in the current implementation — Windows and macOS behave the same from the caller's point of view.

---

## msoffice encrypt

### Syntax

```4d
status:=msoffice encrypt(document; options)
```

| Parameter | Type | Description |
|---|---|---|
| `document` | Blob | The Office Open XML document to encrypt. **Overwritten in place** with the encrypted bytes if the operation succeeds; left unchanged if it fails. |
| `options` | Object | Encryption options (password, secret, and mode) — see below. Optional in the sense that omitting it won't error, but without a password nothing will be encrypted. |
| Result | Object | Status object — see [Description](#description). |

**`options` properties:**

| Property | Type | Description |
|---|---|---|
| `password` | Text | Plain-text password. Highest priority if present. |
| `password_hex` | Text | Password as a hex-encoded byte string. Used only if `password` is absent. |
| `password_uni` | Text | Password as a sequence of `uXXXX` UTF-16 code-unit hex escapes (e.g. `u0041u0042`). Used only if both `password` and `password_hex` are absent. |
| `secret` | Text | Plain-text secret/key value passed through to the encoder alongside the password. Optional. |
| `secret_hex` | Text | Secret as a hex-encoded byte string. Used only if `secret` is absent. |
| `secret_uni` | Text | Secret as a sequence of `uXXXX` UTF-16 code-unit hex escapes. Used only if both `secret` and `secret_hex` are absent. |
| `mode` | Longint | `1` selects AES-256 encryption (the scheme used by Office 2013 and later). Any other value, or omitting `mode`, selects AES-128 (the older/legacy scheme). |

### Description

On success, `document` is replaced in place with the encrypted CFB container, and the returned status object has `success: true` and `isOffice2013` reflecting which AES mode was actually used (`true` for AES-256, `false` for AES-128).

If `document` is already an encrypted CFB container when you call `msoffice encrypt`, the command sets `warning: "already encrypted"` on the status object but still attempts to re-encrypt it — it does not refuse or short-circuit.

If the document couldn't be recognized as a valid Office Open XML container at all (corrupt or unrelated data), the command still attempts encryption, and the status object carries `format: "unknown"` plus an `error` message describing why detection failed. This doesn't by itself guarantee `success: false` — that depends on whether the underlying encoder can still process the bytes.

### Example

From the plugin's own test method (`TEST.4dm`):

```4d
//%attributes = {}
$XLSX:=Folder(fk resources folder).file("unprotected.xlsx").getContent()

$params:=New object("password"; "1234"; "secret"; "abcd")

$status:=msoffice encrypt($XLSX; $params)

If ($status.success)
	
	Folder(fk desktop folder).file("protected.xlsx").setContent($XLSX)
	
End if 
```

Encrypting with AES-256 and no secondary secret:

```4d
$doc:=Document(fk documents folder).file("report.xlsx").getContent()

$options:=New object("password"; "Tr0ub4dor&3"; "mode"; 1)

$status:=msoffice encrypt($doc; $options)

If ($status.success)
	Document(fk documents folder).file("report-protected.xlsx").setContent($doc)
Else
	ALERT("Encryption failed: "+String($status.error))
End if 
```

Supplying the password as a hex-encoded byte string instead of plain text:

```4d
$options:=New object("password_hex"; "313233340041"; "mode"; 0)

$status:=msoffice encrypt($doc; $options)
```

---

## msoffice decrypt

### Syntax

```4d
status:=msoffice decrypt(document; options)
```

| Parameter | Type | Description |
|---|---|---|
| `document` | Blob | The encrypted Office Open XML document. **Overwritten in place** with the decrypted bytes if the operation succeeds; left unchanged if it fails. |
| `options` | Object | Decryption options — see below. Optional in the sense that omitting it won't error, but without the correct password nothing will be decrypted. |
| Result | Object | Status object — see [Description](#description). |

**`options` properties:**

| Property | Type | Description |
|---|---|---|
| `password` | Text | Plain-text password. Highest priority if present. |
| `password_hex` | Text | Password as a hex-encoded byte string. Used only if `password` is absent. |
| `password_uni` | Text | Password as a sequence of `uXXXX` UTF-16 code-unit hex escapes. Used only if both `password` and `password_hex` are absent. |
| `secret` | Text | Plain-text secret/key value passed through to the decoder alongside the password. Optional; must match whatever was used at encryption time. |
| `secret_hex` | Text | Secret as a hex-encoded byte string. Used only if `secret` is absent. |
| `secret_uni` | Text | Secret as a sequence of `uXXXX` UTF-16 code-unit hex escapes. Used only if both `secret` and `secret_hex` are absent. |
| `mode` | — | Ignored by `msoffice decrypt`. The encryption mode is read from the document's own encryption header, not supplied by the caller. |

### Description

On success, `document` is replaced in place with the plain (unencrypted) Office Open XML zip container, and the status object has `success: true`.

If `document` is already a plain, unencrypted zip when you call `msoffice decrypt`, the command sets `warning: "already decrypted"` on the status object but still attempts to decrypt it — it does not refuse or short-circuit. Expect `success: false` in this case, since there's no encryption header to process, but treat `warning` as the reliable signal rather than relying on `success` alone to distinguish "wrong password" from "wasn't encrypted."

A wrong password or secret is reported as `success: false` (typically with an `error` message from the underlying decoder); the input `document` is left unmodified.

### Example

From the plugin's own test method (`TEST.4dm`):

```4d
$XLSX:=Folder(fk resources folder).file("protected.xlsx").getContent()

$params:=New object("password"; "1234"; "secret"; "abcd")

$status:=msoffice decrypt($XLSX; $params)

If ($status.success)
	
	Folder(fk desktop folder).file("unprotected.xlsx").setContent($XLSX)
	
End if 
```

Round-tripping encrypt → decrypt and checking the outcome explicitly:

```4d
$doc:=Document(fk documents folder).file("report-protected.xlsx").getContent()

$options:=New object("password"; "Tr0ub4dor&3"; "secret"; "abcd")

$status:=msoffice decrypt($doc; $options)

Case of 
	: ($status.success)
		Document(fk documents folder).file("report.xlsx").setContent($doc)
	: (Bool($status.warning#Null))
		ALERT("Nothing to decrypt: "+String($status.warning))
	Else 
		ALERT("Decryption failed: "+String($status.error))
End case 
```

---

## Error handling & troubleshooting

- **Omitting all password fields fails silently.** If `options` is missing, or contains none of `password`/`password_hex`/`password_uni`, both commands return `{success: false}` with no `error` field — there's nothing in the status object telling you *why* it failed. Always supply a password if you expect either command to do anything.
- **`warning` doesn't mean the operation was skipped.** `"already encrypted"` and `"already decrypted"` are informational only — the command proceeds to attempt the operation regardless. Check `warning` explicitly if you want to short-circuit on it yourself; don't assume `success: false` implies a warning was the cause, or vice versa.
- **Malformed hex/unicode-hex password or secret values fail the whole call, not just that field.** `password_uni`/`secret_uni` values must be a well-formed run of `uXXXX` 4-hex-digit groups (e.g. `u0041u0042`) with a length that's a multiple of 5 characters; malformed input (wrong length, missing `u` marker, non-hex digits) results in `success: false` with an `error` message rather than the command falling back to a different password field.
- **Only the highest-priority password/secret field present is used** — supplying both `password` and `password_hex` in the same call does not merge or validate them against each other; `password_hex` is simply ignored.
- **`mode` has no effect on `msoffice decrypt`.** The decryption mode is read from the encrypted document's own header. Passing `mode` to decrypt is harmless but does nothing.
- **`document` is unchanged on any failure path.** If `success` is `false` for any reason (bad password, unreadable input, decode error), the `Blob` you passed in is not modified — safe to retry with different options without re-reading the source file.

---

## Quick reference

```4d
// Encrypt
$options:=New object("password"; $pw; "secret"; $secret; "mode"; 1) // AES-256
$status:=msoffice encrypt($doc; $options)
If ($status.success)
	// $doc now holds the encrypted bytes
End if 

// Decrypt
$options:=New object("password"; $pw; "secret"; $secret)
$status:=msoffice decrypt($doc; $options)
If ($status.success)
	// $doc now holds the decrypted bytes
End if 
```
