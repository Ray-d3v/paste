# Paste Help Screenshot Comparison

Reference page: https://pasteapp.io/help/paste-on-mac

## Official Screenshot Signals

- Main history rail: a translucent warm bottom panel over an orange desktop, with a compact toolbar, a selected `Clipboard` pill, colored Pinboard pills, and five visible cards.
- Cards: large rounded tiles with type/time metadata, content previews, app/source labels, and a strong blue selected-card outline.
- Context menu: frosted, rounded macOS-style menu with icon slots, separators, shortcuts, and actions including Open, Paste, Plain Text, Copy, Edit, Rename, Delete, Pin, Quick Look, and Share.
- Edit/search/share states: modal or popover surfaces float above the rail with the same warm translucent material and rounded corners.

Official image refs used:

- History rail: https://framerusercontent.com/images/kNv3GINk8eJUbS2R8m9w3y94wcw.png
- Context menu: https://framerusercontent.com/images/m2RFMBnyM4urf9nKMhwq5Q6tyQ.png
- Edit popover: https://framerusercontent.com/images/noqPGXlWPZJeT1DQOhycykm3I.png
- Search/filter: https://framerusercontent.com/images/mDrKlFuIbkZw7KY8cEJuJ3sQXTM.png
- Pinboard toolbar: https://framerusercontent.com/images/nzUtNNqpMZMt1bz2MNV2iUMKvGQ.png
- Share popover: https://framerusercontent.com/images/yO3qftdOcNTh52ARkVQghEgkxxE.png

## Local Evidence

- Main overlay capture: `artifacts/manual-acceptance/gpui-visual-overlay.png`
- Context menu capture: `artifacts/manual-acceptance/gpui-context-menu.png`

## Current Match

- The default overlay uses a warm frosted bottom panel, compact search/Clipboard/Pinboard toolbar, blue selected-card border, and a five-card rail.
- The comparison rail fixture shows Link, Text, Code, Image, File, and a Favorite marker in the same main screenshot.
- The context menu uses the official action set and order: Open, Paste, Paste as Plain Text, Copy, Edit, Rename, Delete, Pinboard, Quick Look, Share.
- Remaining deliberate differences: generic app/type badges are used instead of copying Paste or app icons, and image/link thumbnails are generated placeholders unless real clipboard payloads include previewable image bytes.
