# Wild Eyez Packing List

This is a tiny standalone site for the Wild Eyez client packing list.

## What it includes

- `index.html` - a simple page with an embedded PDF preview
- `wild_eyez_packing_list.pdf` - the packing list itself

## How to host

The easiest option is GitHub Pages:

1. Create a new GitHub repository with this folder's contents.
2. Turn on GitHub Pages for the repository.
3. Use the published Pages URL inside Wix as an iframe source.

## Wix use

Embed the hosted `index.html` page in an iframe. The page already includes:

- a PDF preview
- a download button
- a direct open button

Recommended Wix page setup:

- Use a real page slug instead of `/blank-1`.
- A cleaner slug would be something like `/packing-list`.
- Point the iframe to the published GitHub Pages URL for this repo.
- If Wix caches the old source, delete the embed element and add a fresh one after changing the slug.

If you want the PDF itself to be easier to read, update the original source export before re-exporting the PDF. The live site can only work with whatever text and logos are already baked into the file.

