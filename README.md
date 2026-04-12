# hwp-to-svg

A pure JavaScript/TypeScript implementation to convert Korean HWP/HWPX (한글) documents to SVG (one SVG per page) or to HWPX format.

## Installation

```bash
npm install hwp-to-svg
```

## CLI Usage

### HWP/HWPX → SVG

```bash
# Convert all pages
hwp2svg document.hwp -o output/

# Convert specific pages
hwp2svg document.hwpx -o output/ -p 0-2
hwp2svg document.hwp -o output/ -p 0,2,4
```

### HWP → HWPX

```bash
hwp2hwpx document.hwp
hwp2hwpx document.hwp -o output.hwpx
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
