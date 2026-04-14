/**
 * Exports a Cytoscape diagram to VSDX format entirely on the client side.
 * Extracts graph data from the Cytoscape instance and builds a VSDX ZIP file using JSZip.
 */
import JSZip from 'jszip'
import {
  CONTENT_TYPES_XML,
  ROOT_RELS_XML,
  DOCUMENT_XML,
  DOCUMENT_RELS_XML,
  PAGES_XML,
  PAGES_RELS_XML,
  MASTERS_XML,
  MASTERS_RELS_XML,
  MASTER1_XML,
  buildAppXml,
  buildCoreXml,
  buildWindowsXml
} from './vsdx-templates'
import { buildPageXml, ImageMap, VsdxEdgeData, VsdxNodeData } from './vsdx-xml-builder'

/**
 * Converts an SVG data URI to a PNG ArrayBuffer using a canvas
 */
function svgToPngBuffer(dataUri: string, width: number, height: number): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      const canvas = document.createElement('canvas')
      const scale = 2 // render at 2x for better quality
      canvas.width = width * scale
      canvas.height = height * scale
      const ctx = canvas.getContext('2d')
      if (!ctx) {
        reject(new Error('Canvas 2d context unavailable'))
        return
      }
      ctx.drawImage(img, 0, 0, width * scale, height * scale)
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error('Canvas toBlob failed'))
          return
        }
        blob.arrayBuffer().then(resolve).catch(reject)
      }, 'image/png')
    }
    img.onerror = () => reject(new Error('Failed to load SVG image'))
    img.src = dataUri
  })
}

/**
 * Converts all unique node background images to PNG buffers.
 * Returns a map of node ID → image key, and a map of image key → ArrayBuffer.
 */
async function convertNodeImages(nodes: VsdxNodeData[]): Promise<{
  imageMap: ImageMap
  imageBuffers: Map<string, ArrayBuffer>
}> {
  const imageMap: ImageMap = new Map()
  const imageBuffers: Map<string, ArrayBuffer> = new Map()
  const dataUriToKey: Map<string, string> = new Map()
  let imageIndex = 0

  for (const node of nodes) {
    if (!node.backgroundImage) {
      continue
    }
    const uri = node.backgroundImage
    if (dataUriToKey.has(uri)) {
      imageMap.set(node.id, dataUriToKey.get(uri) as string)
      continue
    }
    const w = node.backgroundWidth || 24
    const h = node.backgroundHeight || 24
    try {
      const buf = await svgToPngBuffer(uri, w, h)
      imageIndex++
      const key = `image${imageIndex}`
      dataUriToKey.set(uri, key)
      imageMap.set(node.id, key)
      imageBuffers.set(key, buf)
    } catch {
      // Skip nodes where image conversion fails
    }
  }
  return { imageMap, imageBuffers }
}

/**
 * Parses a Cytoscape background-position value, handling both percentage and pixel formats.
 * Returns the pixel offset of the image's leading edge from the node's leading edge.
 *
 * @param rawValue Style value (e.g. "50%" or "10")
 * @param nodeSize Node width/height in pixels
 * @param imgSize Image width/height in pixels
 */
function parseBgPosition(rawValue: string, nodeSize: number, imgSize: number): number {
  if (!rawValue) {
    return (nodeSize - imgSize) / 2
  }
  if (rawValue.includes('%')) {
    const pct = parseFloat(rawValue) / 100
    return pct * (nodeSize - imgSize)
  }
  return parseFloat(rawValue) || (nodeSize - imgSize) / 2
}

/**
 * Extracts node data from a Cytoscape instance, excluding decorator nodes.
 * Invisible helper nodes (e.g. kpi-node) are kept for edge positioning
 * but exported with no fill, no border, and an empty label.
 * Decorator nodes that carry group labels are harvested and their labels
 * applied to the corresponding parent nodes.
 */
function extractNodes(cy: any): VsdxNodeData[] {
  const nodes: VsdxNodeData[] = []

  // Collect group-label decorator labels keyed by their parent node ID
  const decoratorLabels: Map<string, string> = new Map()
  cy.nodes('[?isDecoratorNode]').forEach((deco: any) => {
    const decoData = deco.data()
    const parentId = deco.scratch('decoratorForCytoObject')
    if (parentId && decoData.label && decoData.decoratorName && decoData.decoratorName.includes('group-label')) {
      decoratorLabels.set(parentId, decoData.label)
    }
  })

  // Export visible label-bearing decorator nodes (e.g. edge source/target labels)
  // as text-only shapes so their labels appear in the VSDX output.
  cy.nodes('[?isDecoratorNode]').forEach((deco: any) => {
    const decoData = deco.data()
    if (!decoData.label) {
      return
    }
    // Skip group-label decorators — they're handled via parent node labels
    if (decoData.decoratorName && decoData.decoratorName.includes('group-label')) {
      return
    }
    const pos = deco.position()
    const decoFontSize = parseFloat(deco.style('font-size')) || 12
    const decoFontColor = deco.style('color') || '#000000'
    nodes.push({
      id: decoData.id,
      label: decoData.label,
      x: pos.x,
      y: pos.y,
      width: decoData.label.length * decoFontSize * 0.7,
      height: decoFontSize * 1.5,
      shape: 'rectangle',
      backgroundColor: '#FFFFFF',
      borderColor: '#000000',
      borderWidth: 0,
      borderOpacity: 0,
      backgroundOpacity: 0,
      fontSize: decoFontSize,
      fontColor: decoFontColor,
      isParent: false,
      textHAlign: deco.style('text-halign') || 'center',
      textVAlign: 'center'
    })
  })

  cy.nodes('[!isDecoratorNode]').forEach((node: any) => {
    const pos = node.position()
    const data = node.data()
    const bgImage = node.style('background-image')
    const hasBgImage = bgImage && bgImage !== 'none'

    // Detect invisible layout-helper nodes (zero background opacity or
    // zero overall opacity, no label, no background image). Export them
    // with hidden styling so edges can still connect to them.
    const bgOpacity = parseFloat(node.style('background-opacity'))
    const nodeOpacity = parseFloat(node.style('opacity'))
    const isInvisibleHelper = (bgOpacity === 0 || nodeOpacity === 0) && !hasBgImage

    // For parent nodes with no label, use the decorator label if available
    let nodeLabel = data.label || decoratorLabels.get(data.id) || ''

    // Apply text-transform (e.g., uppercase for NFVO parent nodes)
    const textTransform = node.style('text-transform')
    if (textTransform === 'uppercase' && nodeLabel) {
      nodeLabel = nodeLabel.toUpperCase()
    } else if (textTransform === 'lowercase' && nodeLabel) {
      nodeLabel = nodeLabel.toLowerCase()
    }

    // Truncate label when Cytoscape uses ellipsis text-wrap to match rendered display.
    // Factor 0.45 approximates average character width for variable-width fonts;
    // Cytoscape measures actual glyph widths, so this is a best-effort estimate.
    const textWrap = node.style('text-wrap') || 'none'
    const textMaxWidthPx = parseFloat(node.style('text-max-width')) || 0
    if (textWrap === 'ellipsis' && textMaxWidthPx > 0 && nodeLabel.length > 0) {
      const nodeFontSize = parseFloat(node.style('font-size')) || 12
      const avgCharWidthPx = nodeFontSize * 0.45
      const maxChars = Math.floor(textMaxWidthPx / avgCharWidthPx)
      if (nodeLabel.length > maxChars && maxChars > 3) {
        nodeLabel = nodeLabel.substring(0, maxChars - 3) + '...'
      }
    }

    nodes.push({
      id: data.id,
      label: isInvisibleHelper ? '' : nodeLabel || '',
      x: pos.x,
      y: pos.y,
      width: node.outerWidth() || 80,
      height: node.outerHeight() || 40,
      shape: node.style('shape') || 'rectangle',
      backgroundColor: isInvisibleHelper ? '#FFFFFF' : node.style('background-color') || '#FFFFFF',
      borderColor: node.style('border-color') || '#000000',
      borderWidth: isInvisibleHelper ? 0 : parseFloat(node.style('border-width')) || 1,
      borderOpacity: isInvisibleHelper ? 0 : parseFloat(node.style('border-opacity')) ?? 1,
      backgroundOpacity: bgOpacity,
      fontSize: parseFloat(node.style('font-size')) || 12,
      fontColor: node.style('color') || '#000000',
      parent: data.parent || undefined,
      isParent: node.isParent(),
      backgroundImage: hasBgImage ? bgImage : undefined,
      backgroundWidth: parseFloat(node.style('background-width')) || 0,
      backgroundHeight: parseFloat(node.style('background-height')) || 0,
      backgroundPositionX: parseBgPosition(
        node.style('background-position-x'),
        node.outerWidth() || 80,
        parseFloat(node.style('background-width')) || 0
      ),
      backgroundPositionY: parseBgPosition(
        node.style('background-position-y'),
        node.outerHeight() || 40,
        parseFloat(node.style('background-height')) || 0
      ),
      textHAlign: node.style('text-halign') || 'center',
      textVAlign: node.style('text-valign') || 'bottom',
      backgroundFit: hasBgImage ? node.style('background-fit') || 'none' : 'none',
      textMarginX: parseFloat(node.style('text-margin-x')) || 0,
      textMarginY: parseFloat(node.style('text-margin-y')) || 0
    })
  })

  return nodes
}

/**
 * Extracts edge data from a Cytoscape instance
 */
function extractEdges(cy: any): VsdxEdgeData[] {
  const edges: VsdxEdgeData[] = []

  cy.edges().forEach((edge: any) => {
    const data = edge.data()
    edges.push({
      id: data.id,
      source: data.source,
      target: data.target,
      label: data.label || '',
      lineColor: edge.style('line-color') || '#000000',
      lineWidth: parseFloat(edge.style('width')) || 1,
      lineStyle: edge.style('line-style') || 'solid',
      targetArrowShape: edge.style('target-arrow-shape') || 'none',
      sourceArrowShape: edge.style('source-arrow-shape') || 'none',
      curveStyle: edge.style('curve-style') || 'bezier',
      textMarginY: parseFloat(edge.style('text-margin-y')) || 0
    })
  })

  return edges
}

/**
 * Builds the VSDX ZIP file from Cytoscape graph data and returns it as a Blob
 */
async function buildVsdxBlob(
  nodes: VsdxNodeData[],
  edges: VsdxEdgeData[],
  imageMap: ImageMap,
  imageBuffers: Map<string, ArrayBuffer>
): Promise<Blob> {
  const zip = new JSZip()

  // Build page content first to get page dimensions
  const pageResult = buildPageXml(nodes, edges, imageMap)

  // Root content types — add PNG default if we have images
  let contentTypes = CONTENT_TYPES_XML
  if (imageBuffers.size > 0) {
    contentTypes = contentTypes.replace('</Types>', '  <Default Extension="png" ContentType="image/png"/>\n</Types>')
  }
  zip.file('[Content_Types].xml', contentTypes)

  // Root relationships
  zip.file('_rels/.rels', ROOT_RELS_XML)

  // Document properties
  zip.file('docProps/app.xml', buildAppXml())
  zip.file('docProps/core.xml', buildCoreXml())

  // Visio document
  zip.file('visio/document.xml', DOCUMENT_XML)
  zip.file('visio/_rels/document.xml.rels', DOCUMENT_RELS_XML)

  // Windows (required by Visio desktop)
  zip.file('visio/windows.xml', buildWindowsXml(pageResult.pageWidthInches, pageResult.pageHeightInches))

  // Masters
  zip.file('visio/masters/masters.xml', MASTERS_XML)
  zip.file('visio/masters/_rels/masters.xml.rels', MASTERS_RELS_XML)
  zip.file('visio/masters/master1.xml', MASTER1_XML)

  // Pages
  zip.file('visio/pages/pages.xml', PAGES_XML)
  zip.file('visio/pages/_rels/pages.xml.rels', PAGES_RELS_XML)

  // Page content
  zip.file('visio/pages/page1.xml', pageResult.xml)

  // Page-level relationships for images
  if (pageResult.pageRelsXml) {
    zip.file('visio/pages/_rels/page1.xml.rels', pageResult.pageRelsXml)
  }

  // Media files
  imageBuffers.forEach((buf, key) => {
    zip.file(`visio/media/${key}.png`, buf)
  })

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.ms-visio.drawing',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  })
}

/**
 * Triggers a browser download for the given blob
 */
function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.style.display = 'none'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

/**
 * Exports the Cytoscape diagram to a VSDX file and triggers a browser download.
 *
 * @param cy The Cytoscape instance
 * @param filename The filename for the exported file (without extension)
 */
export async function exportToVsdx(cy: any, filename: string = 'diagram'): Promise<void> {
  const nodes = extractNodes(cy)
  const edges = extractEdges(cy)
  const { imageMap, imageBuffers } = await convertNodeImages(nodes)
  const blob = await buildVsdxBlob(nodes, edges, imageMap, imageBuffers)
  downloadBlob(blob, `${filename}.vsdx`)
}
