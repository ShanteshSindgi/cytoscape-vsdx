# cytoscape-vsdx

Export [Cytoscape.js](https://js.cytoscape.org/) graphs to Visio VSDX format entirely on the client side.

## Installation

```bash
npm install cytoscape-vsdx
```

## Usage

```js
import { exportToVsdx } from 'cytoscape-vsdx';

// Pass your Cytoscape instance and an optional filename
await exportToVsdx(cy, 'my-diagram');
// Downloads my-diagram.vsdx in the browser
```

### What gets exported

- Nodes with labels, colors, borders, and shapes
- Edges with labels, line styles, and arrow shapes
- Group (parent/child) relationships
- Node background images (converted to PNG)
- Text alignment and margins

## API

### `exportToVsdx(cy, filename?)`

Exports the graph and triggers a browser download.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `cy` | Cytoscape instance | — | The Cytoscape.js instance to export |
| `filename` | `string` | `'diagram'` | Name of the downloaded file (without `.vsdx`) |

Returns `Promise<void>`.

### `buildPageXml(nodes, edges, imageMap)`

Lower-level function that builds the Visio page XML from extracted node/edge data. Useful if you need to customize the VSDX generation.

Returns a `PageBuildResult` with the XML string and page dimensions.

## Requirements

- Browser environment (uses Canvas API for image conversion)
- Cytoscape.js 3.x

## License

MIT
