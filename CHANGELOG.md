# Changelog

## 0.1.1

- Renamed the plugin id and repository metadata from `pptx-viewer` / `obsidian-pptx-viewer` to `slide-deck-viewer` / `obsidian-slide-deck-viewer` because the live Community directory already has a `pptx-viewer` plugin.
- Renamed the display name from `PPTX Viewer` to `Slide Deck Viewer`; `powerpoint-viewer` was also already taken in the live Community directory.
- Hardened community checks to validate live Community slug/name collisions, manifest naming/description rules, leaf lifecycle, and unsupported API pitfalls before release.

## 0.1.0

- Initial read-only `.pptx` viewer.
- Added slide navigation, extracted text, speaker notes, media metadata, relationship warnings, and package diagnostics.
- Added fixtures and smoke checks for malformed, encrypted, embedded-object, no-notes, simple, and large decks.

