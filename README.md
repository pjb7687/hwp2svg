# hwp-to-svg

Convert HWP/HWPX (한글) documents to SVG, one SVG per page.

## Installation

```bash
npm install hwp-to-svg
```

## CLI Usage

```bash
# Convert all pages
hwp2svg document.hwp -o output/

# Convert specific pages
hwp2svg document.hwpx -o output/ -p 0-2
hwp2svg document.hwp -o output/ -p 0,2,4
```

## API Usage

```typescript
import { HwpxDocument } from 'hwp-to-svg';

const data = await fs.readFile('document.hwp');
const doc = await HwpxDocument.open(data.buffer);
const pages = doc.renderAllPages(); // string[] of SVG content
await doc.close();
```

## Fonts

The SVG output references fonts by name rather than embedding them. Install the
required Korean fonts system-wide for correct rendering:

```bash
cp fonts/*.TTF ~/.local/share/fonts/
fc-cache -f
```

## License

MIT
