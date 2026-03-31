# json-to-word
An engine to generate docx files from json objects-
[![Gitter](https://badges.gitter.im/json-to-word/community.svg)](https://gitter.im/json-to-word/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

## Notes

- `POST /api/word/create` now supports template-less payloads.  
  When `TemplatePath` is missing/non-absolute, the controller builds a temporary `.docx` with content controls from `ContentControls[*].Title` and then runs the normal `WordService` pipeline.
