# Testing

The public snapshot includes a curated offline test suite. It does not include the complete internal regression suite or production data.

## Run the exported tests

```powershell
python -m pip install -r PropertyManager.App/Backend/requirements.txt
python -m unittest discover -s tests -q
```

The exported suite currently contains **31 tests**:

| Test module | Coverage |
| --- | --- |
| `test_request_policy.py` | Sequential pacing, request reuse, cancellation, rate-limit/challenge handling, transport failures and safe stop behavior |
| `test_workbook_visibility.py` | Windows workbook visibility, staged saves and preservation when file-attribute operations fail |
| `test_campaign_mail.py` | Message-size limits, test-destination rules, attachment checks, HTML/CID handling, submission acknowledgement, duplicate submission and uncertain outcomes |

The workbook visibility tests are Windows-specific. Tests use fake transports and disposable files; they do not need real service credentials.

## Snapshot validation

The prepared public snapshot was additionally checked for:

- A successful Windows .NET build with zero warnings or errors.
- Full-app launcher operation, including a source folder with existing build caches.
- Eight live-operation entry points rejecting execution before external work.
- Separate demo app-data configuration.
- Visual review of the interface and the synthetic documentation screenshots.

The additional preparation checks are not presented as a GitHub CI workflow. No CI status badge or live-service certification is claimed.

## Limits

Passing these checks does not verify a real tenant's API, Outlook profile, SMTP provider, physical printer, Word installation or every Windows configuration. Live entry points are intentionally blocked in this repository. Public tests cover selected boundaries rather than every application feature.
