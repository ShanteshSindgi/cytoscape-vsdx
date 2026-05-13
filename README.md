# cytoscape-vsdx
[![npm version](https://img.shields.io/npm/v/cytoscape-vsdx)](https://www.npmjs.com/package/cytoscape-vsdx)
[![license](https://img.shields.io/npm/l/cytoscape-vsdx)](LICENSE)

Export Cytoscape.js graphs and network diagrams to Microsoft Visio `.vsdx` format entirely on the client side.

## Features

- Export Cytoscape.js graphs to Visio `.vsdx`
- Fully client-side export
- Editable Microsoft Visio output
- Supports nodes, edges, labels, and styles
- Supports compound/group nodes
- Supports node background images
- Preserves edge arrows and line styles
  
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

## Why cytoscape-vsdx?

Most Cytoscape.js export tools only support PNG or SVG exports.

`cytoscape-vsdx` allows exporting Cytoscape.js diagrams into editable Microsoft Visio `.vsdx` files while preserving graph structure and styling.

## Requirements

- Browser environment (uses Canvas API for image conversion)
- Cytoscape.js 3.x

## Keywords

Cytoscape.js, cytoscape-vsdx, VSDX, Microsoft Visio, graph export, network diagrams, Visio exporter, Cytoscape Visio export, JavaScript VSDX library

## License

MIT
