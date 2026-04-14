/**
 * Builds the Visio page XML (page1.xml) from extracted Cytoscape graph data.
 * Maps Cytoscape nodes to Visio shapes and edges to Visio connectors.
 */

// Conversion factor: Cytoscape uses pixels, Visio uses inches. 96 PPI is standard screen DPI.
const PX_TO_INCHES = 1 / 96

/**
 * Data extracted from a Cytoscape node for VSDX export
 */
export interface VsdxNodeData {
  id: string
  label: string
  x: number
  y: number
  width: number
  height: number
  shape: string
  backgroundColor: string
  borderColor: string
  borderWidth: number
  borderOpacity: number
  backgroundOpacity: number
  fontSize: number
  fontColor: string
  parent?: string
  isParent: boolean
  backgroundImage?: string
  backgroundWidth?: number
  backgroundHeight?: number
  backgroundPositionX?: number
  backgroundPositionY?: number
  textHAlign?: string
  textVAlign?: string
  backgroundFit?: string
  textMarginX?: number
  textMarginY?: number
}

/**
 * Mapping from node ID to PNG base64 data for icon images
 */
export type ImageMap = Map<string, string>

/**
 * Data extracted from a Cytoscape edge for VSDX export
 */
export interface VsdxEdgeData {
  id: string
  source: string
  target: string
  label?: string
  lineColor: string
  lineWidth: number
  lineStyle: string
  targetArrowShape: string
  sourceArrowShape: string
  curveStyle: string
  textMarginY: number
}

/**
 * Result of building page XML, includes dimensions needed for windows.xml
 */
export interface PageBuildResult {
  xml: string
  pageWidthInches: number
  pageHeightInches: number
  pageRelsXml: string
}

/**
 * Escapes special XML characters in a string
 */
function escapeXml(str: string): string {
  if (!str) {
    return ''
  }
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/**
 * Converts a CSS color string (hex or rgb) to a Visio color format (#RRGGBB)
 */
function toVisioColor(cssColor: string): string {
  if (!cssColor) {
    return '#000000'
  }

  // Already hex
  if (cssColor.startsWith('#')) {
    return cssColor.length === 4
      ? `#${cssColor[1]}${cssColor[1]}${cssColor[2]}${cssColor[2]}${cssColor[3]}${cssColor[3]}`
      : cssColor
  }

  // rgb(r, g, b) format
  const rgbMatch = cssColor.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/)
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10).toString(16).padStart(2, '0')
    const g = parseInt(rgbMatch[2], 10).toString(16).padStart(2, '0')
    const b = parseInt(rgbMatch[3], 10).toString(16).padStart(2, '0')
    return `#${r}${g}${b}`
  }

  return '#000000'
}

/**
 * Maps a Cytoscape line-style value to a Visio line pattern number
 */
function toVisioLinePattern(lineStyle: string): number {
  switch (lineStyle) {
    case 'dashed':
      return 2
    case 'dotted':
      return 3
    default:
      return 1 // solid
  }
}

/**
 * Maps a Cytoscape arrow shape to a Visio arrow type number
 */
function toVisioArrowType(arrowShape: string): number {
  switch (arrowShape) {
    case 'triangle':
    case 'triangle-backcurve':
    case 'vee':
      return 4 // filled triangle
    case 'circle':
      return 10 // circle
    case 'diamond':
      return 6 // diamond
    case 'none':
    default:
      return 0 // no arrow
  }
}

/**
 * Builds Visio Geometry section for a rectangle with formulas.
 * Visio evaluates F (formula) attributes; V holds the pre-computed value.
 */
function buildRectangleGeometry(w: string, h: string, yOff = 0, yOffFrac = '0'): string {
  const yV = yOff.toFixed(4)
  return `
        <Section N="Geometry" IX="0">
          <Cell N="NoFill" V="0"/>
          <Cell N="NoLine" V="0"/>
          <Row T="MoveTo" IX="1">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="${yV}" F="Height*${yOffFrac}"/>
          </Row>
          <Row T="LineTo" IX="2">
            <Cell N="X" V="${w}" F="Width*1"/>
            <Cell N="Y" V="${yV}" F="Height*${yOffFrac}"/>
          </Row>
          <Row T="LineTo" IX="3">
            <Cell N="X" V="${w}" F="Width*1"/>
            <Cell N="Y" V="${h}" F="Height*1"/>
          </Row>
          <Row T="LineTo" IX="4">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="${h}" F="Height*1"/>
          </Row>
          <Row T="LineTo" IX="5">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="${yV}" F="Height*${yOffFrac}"/>
          </Row>
        </Section>`
}

/**
 * Builds Visio Geometry section for an ellipse with formulas
 */
function buildEllipseGeometry(w: string, h: string, yOff = 0, yOffFrac = '0'): string {
  const totalH = parseFloat(h)
  const centerY = (yOff + totalH) / 2
  const halfW = (parseFloat(w) / 2).toFixed(4)
  const centerYFrac = ((parseFloat(yOffFrac) + 1) / 2).toFixed(6)
  const topFrac = '1'
  const rightFrac = centerYFrac
  return `
        <Section N="Geometry" IX="0">
          <Cell N="NoFill" V="0"/>
          <Cell N="NoLine" V="0"/>
          <Row T="Ellipse" IX="1">
            <Cell N="X" V="${halfW}" F="Width*0.5"/>
            <Cell N="Y" V="${centerY.toFixed(4)}" F="Height*${centerYFrac}"/>
            <Cell N="A" V="${w}" F="Width*1"/>
            <Cell N="B" V="${centerY.toFixed(4)}" F="Height*${rightFrac}"/>
            <Cell N="C" V="${halfW}" F="Width*0.5"/>
            <Cell N="D" V="${totalH.toFixed(4)}" F="Height*${topFrac}"/>
          </Row>
        </Section>`
}

/**
 * Builds Visio Geometry section for a diamond with formulas
 */
function buildDiamondGeometry(w: string, h: string, yOff = 0, yOffFrac = '0'): string {
  const totalH = parseFloat(h)
  const midY = (yOff + totalH) / 2
  const halfW = (parseFloat(w) / 2).toFixed(4)
  const midFrac = ((parseFloat(yOffFrac) + 1) / 2).toFixed(6)
  return `
        <Section N="Geometry" IX="0">
          <Cell N="NoFill" V="0"/>
          <Cell N="NoLine" V="0"/>
          <Row T="MoveTo" IX="1">
            <Cell N="X" V="${halfW}" F="Width*0.5"/>
            <Cell N="Y" V="${yOff.toFixed(4)}" F="Height*${yOffFrac}"/>
          </Row>
          <Row T="LineTo" IX="2">
            <Cell N="X" V="${w}" F="Width*1"/>
            <Cell N="Y" V="${midY.toFixed(4)}" F="Height*${midFrac}"/>
          </Row>
          <Row T="LineTo" IX="3">
            <Cell N="X" V="${halfW}" F="Width*0.5"/>
            <Cell N="Y" V="${totalH.toFixed(4)}" F="Height*1"/>
          </Row>
          <Row T="LineTo" IX="4">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="${midY.toFixed(4)}" F="Height*${midFrac}"/>
          </Row>
          <Row T="LineTo" IX="5">
            <Cell N="X" V="${halfW}" F="Width*0.5"/>
            <Cell N="Y" V="${yOff.toFixed(4)}" F="Height*${yOffFrac}"/>
          </Row>
        </Section>`
}

/**
 * Returns the geometry section for a given Cytoscape shape
 */
function buildGeometryForShape(
  shape: string,
  widthInches: number,
  totalHeightInches: number,
  yOffsetInches: number
): string {
  const w = widthInches.toFixed(4)
  const h = totalHeightInches.toFixed(4)
  const yOffFrac = (yOffsetInches / totalHeightInches).toFixed(6)
  switch (shape) {
    case 'ellipse':
      return buildEllipseGeometry(w, h, yOffsetInches, yOffFrac)
    case 'diamond':
      return buildDiamondGeometry(w, h, yOffsetInches, yOffFrac)
    case 'rectangle':
    case 'roundrectangle':
    case 'round-rectangle':
    default:
      return buildRectangleGeometry(w, h, yOffsetInches, yOffFrac)
  }
}

/**
 * Builds a Visio Shape XML element for a Cytoscape node
 */
function buildShapeXml(node: VsdxNodeData, shapeId: number, pageHeightInches: number): string {
  const widthInches = node.width * PX_TO_INCHES
  const heightInches = node.height * PX_TO_INCHES
  // Visio Y axis is bottom-up, Cytoscape is top-down
  const pinX = node.x * PX_TO_INCHES
  const pinY = pageHeightInches - node.y * PX_TO_INCHES
  const fillColor = toVisioColor(node.backgroundColor)
  const lineColor = toVisioColor(node.borderColor)
  const fontColor = toVisioColor(node.fontColor)
  const lineWeight = (node.borderWidth || 1) * PX_TO_INCHES
  const fontSize = ((node.fontSize || 12) / 72).toFixed(4) // points to inches
  const labelText = escapeXml(node.label || '')
  const isRounded = node.shape === 'roundrectangle' || node.shape === 'round-rectangle'
  const rounding = isRounded ? `\n        <Cell N="Rounding" V="0.1" U="IN"/>` : ''
  // Visio LineColorTrans: 0 = opaque, 1 = fully transparent (fraction, not integer)
  const lineTransparency = (1 - (node.borderOpacity ?? 1)).toFixed(4)
  // Invisible nodes: either explicit helper nodes or icon-only nodes (background-opacity: 0)
  const isInvisible =
    (node.borderWidth === 0 && node.borderOpacity === 0 && !node.label) || node.backgroundOpacity === 0
  const fillPattern = isInvisible ? 0 : 1
  const linePattern = isInvisible ? 0 : 1

  // Parent/compound nodes: label at top, no text zone below
  // Map Cytoscape text-halign to Visio HorzAlign: left=0, center=1, right=2
  const parentHorzAlign = node.textHAlign === 'left' ? 0 : node.textHAlign === 'right' ? 2 : 1
  if (node.isParent) {
    const geometry = buildGeometryForShape(node.shape, widthInches, heightInches, 0)
    // Text block at top of compound shape — use a fixed height label strip
    const txtBlockH = 0.3
    const txtPinYParent = heightInches - txtBlockH / 2
    const txtPinYFracParent = (txtPinYParent / heightInches).toFixed(6)
    const txtHFracParent = (txtBlockH / heightInches).toFixed(6)
    return `
      <Shape ID="${shapeId}" Type="Shape" NameU="Sheet.${shapeId}">
        <Cell N="PinX" V="${pinX.toFixed(4)}" U="IN"/>
        <Cell N="PinY" V="${pinY.toFixed(4)}" U="IN"/>
        <Cell N="Width" V="${widthInches.toFixed(4)}" U="IN"/>
        <Cell N="Height" V="${heightInches.toFixed(4)}" U="IN"/>
        <Cell N="LocPinX" V="${(widthInches / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="LocPinY" V="${(heightInches / 2).toFixed(4)}" U="IN" F="Height*0.5"/>
        <Cell N="FillForegnd" V="${fillColor}"/>
        <Cell N="FillBkgnd" V="${fillColor}"/>
        <Cell N="FillPattern" V="${fillPattern}"/>
        <Cell N="LineColor" V="${lineColor}"/>
        <Cell N="LineWeight" V="${lineWeight.toFixed(4)}" U="IN"/>
        <Cell N="LinePattern" V="${linePattern}"/>
        <Cell N="LineColorTrans" V="${lineTransparency}"/>${rounding}
        <Cell N="TxtPinX" V="${(widthInches / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="TxtPinY" V="${txtPinYParent.toFixed(4)}" U="IN" F="Height*${txtPinYFracParent}"/>
        <Cell N="TxtLocPinX" V="${(widthInches / 2).toFixed(4)}" U="IN" F="TxtWidth*0.5"/>
        <Cell N="TxtLocPinY" V="${(txtBlockH / 2).toFixed(4)}" U="IN" F="TxtHeight*0.5"/>
        <Cell N="TxtWidth" V="${widthInches.toFixed(4)}" U="IN" F="Width*1"/>
        <Cell N="TxtHeight" V="${txtBlockH.toFixed(4)}" U="IN" F="Height*${txtHFracParent}"/>
        <Cell N="VerticalAlign" V="1"/>
        <Section N="Character">
          <Row IX="0">
            <Cell N="Font" V="Calibri"/>
            <Cell N="Color" V="${fontColor}"/>
            <Cell N="Size" V="${fontSize}" U="IN"/>
            <Cell N="Style" V="1"/>
          </Row>
        </Section>
        <Section N="Paragraph">
          <Row IX="0">
            <Cell N="HorzAlign" V="${parentHorzAlign}"/>
          </Row>
        </Section>${geometry}
        <Text><cp IX="0"/><pp IX="0"/>${labelText}</Text>
      </Shape>`
  }

  // Child/leaf nodes: text placement depends on text-valign and text-halign
  // 'center' → text inside, vertically centered
  // 'top'    → text inside, aligned to top
  // 'bottom' → text in a dedicated zone below the box
  // When background-fit is 'contain' or 'cover', the icon fills the entire node,
  // so text must go below regardless of text-valign to avoid overlapping the icon.
  const iconFillsNode = node.backgroundFit === 'contain' || node.backgroundFit === 'cover'
  const effectiveVAlign = iconFillsNode && node.backgroundImage ? 'bottom' : node.textVAlign || 'bottom'

  // Estimate text dimensions early — needed for side-text detection
  const fontSizeInches = (node.fontSize || 12) / 72
  const estimatedTextWidth = (node.label || '').length * fontSizeInches * 0.6

  // When text-halign is left or right and the label is wider than the node,
  // position text beside the node (to its left or right) instead of inside
  // or below it. This handles port nodes and other small shapes where
  // Cytoscape displays labels adjacent to the shape.
  const textBesideNode =
    (node.textHAlign === 'left' || node.textHAlign === 'right') && estimatedTextWidth > widthInches * 0.8

  const textInside = !textBesideNode && (effectiveVAlign === 'center' || effectiveVAlign === 'top')
  const txtZoneHeight = textInside || textBesideNode ? 0 : 0.3
  const totalHeight = heightInches + txtZoneHeight
  const geometry = buildGeometryForShape(node.shape, widthInches, totalHeight, txtZoneHeight)
  // LocPinY places the visual box center at (pinX, pinY) on the page
  const locPinY = txtZoneHeight + heightInches / 2
  const locPinYFrac = (locPinY / totalHeight).toFixed(6)
  const txtWidth = textBesideNode ? estimatedTextWidth + 0.2 : Math.max(widthInches, estimatedTextWidth + 0.2)

  // Text block positioning
  let txtPinY: number
  let txtPinYFrac: string
  let txtHeight: number
  let txtHeightFrac: string
  let verticalAlign: number

  // Horizontal text positioning — default is centered on the shape
  let txtPinXV = (widthInches / 2).toFixed(4)
  let txtPinXF = 'Width*0.5'
  let txtLocPinXV = (txtWidth / 2).toFixed(4)
  let txtLocPinXF = 'TxtWidth*0.5'
  let horzAlign = 1

  if (textBesideNode) {
    // Text block positioned to the left or right of the node
    const marginXInches = (node.textMarginX || 0) * PX_TO_INCHES
    if (node.textHAlign === 'right') {
      // Anchor text at right edge, extends rightward
      const anchorX = widthInches + marginXInches
      txtPinXV = anchorX.toFixed(4)
      txtPinXF = `Width*${(widthInches > 0 ? anchorX / widthInches : 1).toFixed(6)}`
      txtLocPinXV = '0.0000'
      txtLocPinXF = 'TxtWidth*0'
      horzAlign = 0 // left-aligned so text flows rightward
    } else {
      // Anchor text at left edge, extends leftward
      const anchorX = marginXInches
      txtPinXV = anchorX.toFixed(4)
      txtPinXF = `Width*${(widthInches > 0 ? anchorX / widthInches : 0).toFixed(6)}`
      txtLocPinXV = txtWidth.toFixed(4)
      txtLocPinXF = 'TxtWidth*1'
      horzAlign = 2 // right-aligned so text flows leftward
    }
    // Vertically center text on the node
    const marginYInches = (node.textMarginY || 0) * PX_TO_INCHES
    txtPinY = totalHeight / 2 - marginYInches
    txtPinYFrac = (txtPinY / totalHeight).toFixed(6)
    txtHeight = Math.max(fontSizeInches * 2, totalHeight)
    txtHeightFrac = '' // no formula — height is independent of shape
    verticalAlign = 1
  } else if (effectiveVAlign === 'top') {
    // Text inside the shape, aligned to top
    txtPinY = heightInches / 2
    txtPinYFrac = '0.5'
    txtHeight = heightInches
    txtHeightFrac = '1'
    verticalAlign = 0 // top
  } else if (effectiveVAlign === 'center') {
    // Text inside the shape, vertically centered
    txtPinY = heightInches / 2
    txtPinYFrac = '0.5'
    txtHeight = heightInches
    txtHeightFrac = '1'
    verticalAlign = 1 // middle
  } else {
    // Text in a zone below the visual box
    txtPinY = txtZoneHeight / 2
    txtPinYFrac = (txtPinY / totalHeight).toFixed(6)
    txtHeight = txtZoneHeight
    txtHeightFrac = (txtZoneHeight / totalHeight).toFixed(6)
    verticalAlign = 1
  }

  // Build TxtHeight cell — include F formula only when height tracks shape Height
  const txtHeightCell = txtHeightFrac
    ? `<Cell N="TxtHeight" V="${txtHeight.toFixed(4)}" U="IN" F="Height*${txtHeightFrac}"/>`
    : `<Cell N="TxtHeight" V="${txtHeight.toFixed(4)}" U="IN"/>`

  return `
      <Shape ID="${shapeId}" Type="Shape" NameU="Sheet.${shapeId}">
        <Cell N="PinX" V="${pinX.toFixed(4)}" U="IN"/>
        <Cell N="PinY" V="${pinY.toFixed(4)}" U="IN"/>
        <Cell N="Width" V="${widthInches.toFixed(4)}" U="IN"/>
        <Cell N="Height" V="${totalHeight.toFixed(4)}" U="IN"/>
        <Cell N="LocPinX" V="${(widthInches / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="LocPinY" V="${locPinY.toFixed(4)}" U="IN" F="Height*${locPinYFrac}"/>
        <Cell N="FillForegnd" V="${fillColor}"/>
        <Cell N="FillBkgnd" V="${fillColor}"/>
        <Cell N="FillPattern" V="${fillPattern}"/>
        <Cell N="LineColor" V="${lineColor}"/>
        <Cell N="LineWeight" V="${lineWeight.toFixed(4)}" U="IN"/>
        <Cell N="LinePattern" V="${linePattern}"/>
        <Cell N="LineColorTrans" V="${lineTransparency}"/>${rounding}
        <Cell N="TxtPinX" V="${txtPinXV}" U="IN" F="${txtPinXF}"/>
        <Cell N="TxtPinY" V="${txtPinY.toFixed(4)}" U="IN" F="Height*${txtPinYFrac}"/>
        <Cell N="TxtLocPinX" V="${txtLocPinXV}" U="IN" F="${txtLocPinXF}"/>
        <Cell N="TxtLocPinY" V="${(txtHeight / 2).toFixed(4)}" U="IN" F="TxtHeight*0.5"/>
        <Cell N="TxtWidth" V="${txtWidth.toFixed(4)}" U="IN"/>
        ${txtHeightCell}
        <Cell N="VerticalAlign" V="${verticalAlign}"/>
        <Section N="Character">
          <Row IX="0">
            <Cell N="Font" V="Calibri"/>
            <Cell N="Color" V="${fontColor}"/>
            <Cell N="Size" V="${fontSize}" U="IN"/>
            <Cell N="Style" V="0"/>
          </Row>
        </Section>
        <Section N="Paragraph">
          <Row IX="0">
            <Cell N="HorzAlign" V="${horzAlign}"/>
            <Cell N="SpLine" V="-1.2"/>
          </Row>
        </Section>
        <Section N="Connection" IX="0">
          <Row IX="0">
            <Cell N="X" V="${(widthInches / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
            <Cell N="Y" V="${locPinY.toFixed(4)}" U="IN" F="Height*${locPinYFrac}"/>
          </Row>
        </Section>${geometry}
        <Text><cp IX="0"/><pp IX="0"/>${labelText}</Text>
      </Shape>`
}

/**
 * Computes where a line from a rectangle's center to an external point
 * intersects the rectangle's perimeter. All values in inches.
 */
function rectPerimeterPoint(
  cx: number,
  cy: number,
  halfW: number,
  halfH: number,
  tx: number,
  ty: number
): { x: number; y: number } {
  const dx = tx - cx
  const dy = ty - cy
  if (dx === 0 && dy === 0) {
    return { x: cx + halfW, y: cy }
  }
  const absDx = Math.abs(dx)
  const absDy = Math.abs(dy)
  const scale = absDx * halfH > absDy * halfW ? halfW / absDx : halfH / absDy
  return { x: cx + dx * scale, y: cy + dy * scale }
}

/**
 * Builds a Visio connector (Shape) XML for a Cytoscape edge
 */
function buildConnectorXml(
  edge: VsdxEdgeData,
  shapeId: number,
  sourceNode: VsdxNodeData,
  targetNode: VsdxNodeData,
  pageHeightInches: number,
  sourceShapeId: number,
  targetShapeId: number
): string {
  // Node centers in Visio coordinates
  const srcCx = sourceNode.x * PX_TO_INCHES
  const srcCy = pageHeightInches - sourceNode.y * PX_TO_INCHES
  const tgtCx = targetNode.x * PX_TO_INCHES
  const tgtCy = pageHeightInches - targetNode.y * PX_TO_INCHES

  // Compute perimeter points so V values connect to box edges, not centers.
  // _WALKGLUE handles this dynamically after drag, but Visio doesn't always
  // evaluate formulas on initial open — the V fallback must be correct.
  const srcHalfW = (sourceNode.width * PX_TO_INCHES) / 2
  const srcHalfH = (sourceNode.height * PX_TO_INCHES) / 2
  const tgtHalfW = (targetNode.width * PX_TO_INCHES) / 2
  const tgtHalfH = (targetNode.height * PX_TO_INCHES) / 2

  const beginPt = rectPerimeterPoint(srcCx, srcCy, srcHalfW, srcHalfH, tgtCx, tgtCy)
  const endPt = rectPerimeterPoint(tgtCx, tgtCy, tgtHalfW, tgtHalfH, srcCx, srcCy)

  const beginX = beginPt.x
  const beginY = beginPt.y
  const endX = endPt.x
  const endY = endPt.y
  const lineColor = toVisioColor(edge.lineColor)
  const lineWeight = (edge.lineWidth || 1) * PX_TO_INCHES
  const linePattern = toVisioLinePattern(edge.lineStyle)
  const beginArrow = toVisioArrowType(edge.sourceArrowShape)
  const endArrow = toVisioArrowType(edge.targetArrowShape)
  const labelText = escapeXml(edge.label || '')

  // 1-D connector positioning
  const connW = Math.sqrt(Math.pow(endX - beginX, 2) + Math.pow(endY - beginY, 2))
  const midX = (beginX + endX) / 2
  const midY = (beginY + endY) / 2
  const angle = Math.atan2(endY - beginY, endX - beginX)

  // Keep text always horizontal (angle 0) regardless of connector direction.
  // TxtAngle is relative to the shape's own rotation, so negate it.
  const txtAngle = -angle

  // Position the text block above the connector line.
  // In the connector's local coordinate system, Y=0 is the line itself.
  // A positive TxtPinY offset moves text above the line.
  const textMarginInches = Math.abs(edge.textMarginY) * PX_TO_INCHES
  const rawTxtOffset = textMarginInches > 0 ? textMarginInches : labelText ? 0.15 : 0
  // When the connector points right-to-left (|angle| > π/2), the local Y axis
  // is inverted relative to the page. Negate the offset so labels stay visually
  // above the line regardless of connector direction.
  const txtOffset = Math.cos(angle) >= 0 ? rawTxtOffset : -rawTxtOffset

  // Visio route style for different Cytoscape curve-styles
  // 16 = right-angle, 1 = straight, 6 = simple (center-to-center)
  const routeStyle = edge.curveStyle === 'taxi' ? 16 : 1

  // _WALKGLUE makes endpoints walk to shape perimeter in Visio.
  // BegTrigger/EndTrigger recalculate when connected shape moves.
  // The V attribute holds the static fallback position for tools that can't evaluate
  // Visio-proprietary formulas (e.g. Lucidchart).
  const srcSheet = `Sheet.${sourceShapeId}`
  const tgtSheet = `Sheet.${targetShapeId}`

  return `
      <Shape ID="${shapeId}" Type="Shape" NameU="Connector.${shapeId}" LineStyle="0" FillStyle="0" TextStyle="0">
        <Cell N="PinX" V="${midX.toFixed(4)}" U="IN" F="(BeginX+EndX)/2"/>
        <Cell N="PinY" V="${midY.toFixed(4)}" U="IN" F="(BeginY+EndY)/2"/>
        <Cell N="Width" V="${connW.toFixed(
          4
        )}" U="IN" F="SQRT((BeginX-EndX)*(BeginX-EndX)+(BeginY-EndY)*(BeginY-EndY))"/>
        <Cell N="Height" V="0" U="IN"/>
        <Cell N="LocPinX" V="${(connW / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="LocPinY" V="0" U="IN" F="Height*0.5"/>
        <Cell N="Angle" V="${angle.toFixed(6)}" F="ATAN2(EndY-BeginY,EndX-BeginX)"/>
        <Cell N="TxtAngle" V="${txtAngle.toFixed(6)}"/>
        <Cell N="TxtPinX" V="${(connW / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="TxtPinY" V="${txtOffset.toFixed(4)}" U="IN"/>
        <Cell N="TxtLocPinX" V="${(connW / 2).toFixed(4)}" U="IN" F="TxtWidth*0.5"/>
        <Cell N="TxtLocPinY" V="0" U="IN"/>
        <Cell N="TxtWidth" V="${connW.toFixed(4)}" U="IN" F="Width*1"/>
        <Cell N="TxtHeight" V="0.25" U="IN"/>
        <Cell N="BeginX" V="${beginX.toFixed(
          4
        )}" U="IN" F="_WALKGLUE(BegTrigger,${srcSheet}!EventXFMod,${srcSheet}!EventXFMod)"/>
        <Cell N="BeginY" V="${beginY.toFixed(
          4
        )}" U="IN" F="_WALKGLUE(BegTrigger,${srcSheet}!EventXFMod,${srcSheet}!EventXFMod)"/>
        <Cell N="EndX" V="${endX.toFixed(
          4
        )}" U="IN" F="_WALKGLUE(EndTrigger,${tgtSheet}!EventXFMod,${tgtSheet}!EventXFMod)"/>
        <Cell N="EndY" V="${endY.toFixed(
          4
        )}" U="IN" F="_WALKGLUE(EndTrigger,${tgtSheet}!EventXFMod,${tgtSheet}!EventXFMod)"/>
        <Cell N="BegTrigger" V="2" F="_XFTRIGGER(${srcSheet}!EventXFMod)"/>
        <Cell N="EndTrigger" V="2" F="_XFTRIGGER(${tgtSheet}!EventXFMod)"/>
        <Cell N="ObjType" V="2"/>
        <Cell N="ShapeRouteStyle" V="${routeStyle}"/>
        <Cell N="LineColor" V="${lineColor}"/>
        <Cell N="LineWeight" V="${lineWeight.toFixed(4)}" U="IN"/>
        <Cell N="LinePattern" V="${linePattern}"/>
        <Cell N="BeginArrow" V="${beginArrow}"/>
        <Cell N="EndArrow" V="${endArrow}"/>
        <Cell N="FillPattern" V="0"/>
        <Section N="Character">
          <Row IX="0">
            <Cell N="Font" V="Calibri"/>
            <Cell N="Color" V="${lineColor}"/>
            <Cell N="Size" V="0.1389" U="IN"/>
            <Cell N="Style" V="0"/>
          </Row>
        </Section>
        <Section N="Paragraph">
          <Row IX="0">
            <Cell N="HorzAlign" V="1"/>
          </Row>
        </Section>
        <Section N="Geometry" IX="0">
          <Cell N="NoFill" V="1"/>
          <Cell N="NoLine" V="0"/>
          <Row T="MoveTo" IX="1">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="0"/>
          </Row>
          <Row T="LineTo" IX="2">
            <Cell N="X" V="${connW.toFixed(4)}" F="Width*1"/>
            <Cell N="Y" V="0"/>
          </Row>
        </Section>
        <Text><cp IX="0"/><pp IX="0"/>${labelText}</Text>
      </Shape>`
}

/**
 * Builds the complete page1.xml content from node and edge data.
 * Returns the XML string along with page dimensions for use in windows.xml.
 */
export function buildPageXml(nodes: VsdxNodeData[], edges: VsdxEdgeData[], imageMap?: ImageMap): PageBuildResult {
  // Calculate page dimensions from bounding box of all nodes
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity

  nodes.forEach((node) => {
    const halfW = node.width / 2
    const halfH = node.height / 2
    minX = Math.min(minX, node.x - halfW)
    minY = Math.min(minY, node.y - halfH)
    maxX = Math.max(maxX, node.x + halfW)
    maxY = Math.max(maxY, node.y + halfH)
  })

  // Add padding (in pixels)
  const padding = 100
  minX -= padding
  minY -= padding
  maxX += padding
  maxY += padding

  const pageWidthInches = (maxX - minX) * PX_TO_INCHES
  const pageHeightInches = (maxY - minY) * PX_TO_INCHES

  // Offset all node positions so the minimum is at the padding origin
  const offsetNodes: VsdxNodeData[] = nodes.map((n) => ({
    ...n,
    x: n.x - minX,
    y: n.y - minY
  }))

  // Build a lookup map for source/target resolution
  const nodeMap: Map<string, VsdxNodeData> = new Map()
  offsetNodes.forEach((n) => nodeMap.set(n.id, n))

  let shapeId = 1
  const shapes: string[] = []

  // Track node ID → Visio shape ID for connector binding
  const nodeIdToShapeId: Map<string, number> = new Map()

  // Pre-compute which nodes will get a text overlay (icon + text inside box).
  // For these nodes, suppress text in the base shape to prevent doubling —
  // the transparent icon PNG lets shape text bleed through behind the overlay.
  // Skip nodes where icon fills the entire node (background-fit: contain/cover)
  // — those already have text placed below the box, not overlapping the icon.
  const nodesWithTextOverlay = new Set<string>()
  if (imageMap) {
    offsetNodes.forEach((node) => {
      const iconFills = node.backgroundFit === 'contain' || node.backgroundFit === 'cover'
      if (
        imageMap.get(node.id) &&
        !node.isParent &&
        !iconFills &&
        node.label &&
        (node.textVAlign === 'top' || node.textVAlign === 'center')
      ) {
        nodesWithTextOverlay.add(node.id)
      }
    })
  }

  // Build node shapes
  offsetNodes.forEach((node) => {
    nodeIdToShapeId.set(node.id, shapeId)
    const shapeNode = nodesWithTextOverlay.has(node.id) ? { ...node, label: '' } : node
    shapes.push(buildShapeXml(shapeNode, shapeId, pageHeightInches))
    shapeId++
  })

  // Build edge connectors and collect connection info
  const connects: string[] = []
  edges.forEach((edge) => {
    const sourceNode = nodeMap.get(edge.source)
    const targetNode = nodeMap.get(edge.target)
    if (sourceNode && targetNode) {
      const connectorId = shapeId
      const sourceShapeId = nodeIdToShapeId.get(edge.source)
      const targetShapeId = nodeIdToShapeId.get(edge.target)
      shapes.push(
        buildConnectorXml(edge, shapeId, sourceNode, targetNode, pageHeightInches, sourceShapeId, targetShapeId)
      )
      if (sourceShapeId !== undefined) {
        connects.push(
          `    <Connect FromSheet="${connectorId}" FromCell="BeginX" FromPart="9" ToSheet="${sourceShapeId}" ToCell="Connections.X1" ToPart="3"/>`
        )
      }
      if (targetShapeId !== undefined) {
        connects.push(
          `    <Connect FromSheet="${connectorId}" FromCell="EndX" FromPart="12" ToSheet="${targetShapeId}" ToCell="Connections.X1" ToPart="3"/>`
        )
      }
      shapeId++
    }
  })

  const connectsSection = connects.length > 0 ? `\n  <Connects>\n${connects.join('\n')}\n  </Connects>` : ''

  // Build icon overlay shapes for nodes with background images
  const imageRels: { relId: string; imageKey: string }[] = []
  let relIndex = 1
  if (imageMap) {
    offsetNodes.forEach((node) => {
      const imageKey = imageMap.get(node.id)
      if (!imageKey) {
        return
      }
      // Skip icon overlays for parent/compound nodes — their background images
      // are decorative header strips that would cover the label text.
      if (node.isParent) {
        return
      }
      const imgW = (node.backgroundWidth || 24) * PX_TO_INCHES
      const imgH = (node.backgroundHeight || 24) * PX_TO_INCHES
      const nodeW = node.width * PX_TO_INCHES
      const nodeH = node.height * PX_TO_INCHES
      // Position: background-position is from top-left of node
      const posX = (node.backgroundPositionX || 0) * PX_TO_INCHES
      const posY = (node.backgroundPositionY || 0) * PX_TO_INCHES
      // Convert to Visio coordinates (center-based, Y flipped)
      const nodeLeft = node.x * PX_TO_INCHES - nodeW / 2
      const nodeTop = pageHeightInches - (node.y * PX_TO_INCHES - nodeH / 2)
      const iconPinX = nodeLeft + posX + imgW / 2
      const iconPinY = nodeTop - posY - imgH / 2

      const relId = `rId${relIndex}`
      // Deduplicate: only add one rel per unique image key
      if (!imageRels.some((r) => r.imageKey === imageKey)) {
        imageRels.push({ relId, imageKey })
        relIndex++
      }
      const actualRelId = (imageRels.find((r) => r.imageKey === imageKey) as { relId: string; imageKey: string }).relId

      shapes.push(`
      <Shape ID="${shapeId}" Type="Foreign" NameU="Icon.${shapeId}">
        <Cell N="PinX" V="${iconPinX.toFixed(4)}" U="IN"/>
        <Cell N="PinY" V="${iconPinY.toFixed(4)}" U="IN"/>
        <Cell N="Width" V="${imgW.toFixed(4)}" U="IN"/>
        <Cell N="Height" V="${imgH.toFixed(4)}" U="IN"/>
        <Cell N="LocPinX" V="${(imgW / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="LocPinY" V="${(imgH / 2).toFixed(4)}" U="IN" F="Height*0.5"/>
        <Cell N="ImgWidth" V="${imgW.toFixed(4)}" U="IN" F="Width*1"/>
        <Cell N="ImgHeight" V="${imgH.toFixed(4)}" U="IN" F="Height*1"/>
        <Cell N="ImgOffsetX" V="0" U="IN"/>
        <Cell N="ImgOffsetY" V="0" U="IN"/>
        <Cell N="FillForegnd" V="#FFFFFF"/>
        <Cell N="FillPattern" V="1"/>
        <Cell N="FillForegndTrans" V="1"/>
        <Cell N="LinePattern" V="0"/>
        <Cell N="SelectMode" V="1"/>
        <Section N="Geometry" IX="0">
          <Cell N="NoFill" V="0"/>
          <Cell N="NoLine" V="1"/>
          <Row T="MoveTo" IX="1">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="0" F="Height*0"/>
          </Row>
          <Row T="LineTo" IX="2">
            <Cell N="X" V="${imgW.toFixed(4)}" F="Width*1"/>
            <Cell N="Y" V="0" F="Height*0"/>
          </Row>
          <Row T="LineTo" IX="3">
            <Cell N="X" V="${imgW.toFixed(4)}" F="Width*1"/>
            <Cell N="Y" V="${imgH.toFixed(4)}" F="Height*1"/>
          </Row>
          <Row T="LineTo" IX="4">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="${imgH.toFixed(4)}" F="Height*1"/>
          </Row>
          <Row T="LineTo" IX="5">
            <Cell N="X" V="0" F="Width*0"/>
            <Cell N="Y" V="0" F="Height*0"/>
          </Row>
        </Section>
        <ForeignData ForeignType="Bitmap" CompressionType="PNG">
          <Rel r:id="${actualRelId}"/>
        </ForeignData>
      </Shape>`)
      shapeId++

      // For nodes with text inside the box (text-valign: top/center), the icon
      // Foreign shape covers the label. Add a text-only overlay shape on top.
      // Skip nodes where icon fills the entire node (contain/cover) — those
      // have text below the box already, not overlapping the icon.
      const iconFillsOverlay = node.backgroundFit === 'contain' || node.backgroundFit === 'cover'
      if (node.label && !iconFillsOverlay && (node.textVAlign === 'top' || node.textVAlign === 'center')) {
        const txtNodeW = node.width * PX_TO_INCHES
        const txtNodeH = node.height * PX_TO_INCHES
        const txtFontSizeInches = (node.fontSize || 12) / 72
        const txtEstTextW = (node.label || '').length * txtFontSizeInches * 0.6
        const txtFontSize = txtFontSizeInches.toFixed(4)
        const txtFontColor = toVisioColor(node.fontColor)
        const txtVAlign = node.textVAlign === 'top' ? 0 : 1

        // When text-halign is left/right and the label is wider than the node,
        // position the overlay shape beside the node so text is not constrained
        // to the tiny node width. This fixes port labels in RAN diagrams.
        const overlayBeside =
          (node.textHAlign === 'left' || node.textHAlign === 'right') && txtEstTextW > txtNodeW * 0.8
        let overlayPinX = node.x * PX_TO_INCHES
        let overlayW = txtNodeW
        let overlayH = txtNodeH
        let overlayHAlign = 1
        if (overlayBeside) {
          overlayW = txtEstTextW + 0.2
          overlayH = Math.max(txtFontSizeInches * 2, txtNodeH)
          const marginX = (node.textMarginX || 0) * PX_TO_INCHES
          if (node.textHAlign === 'right') {
            overlayPinX = node.x * PX_TO_INCHES + txtNodeW / 2 + marginX + overlayW / 2
            overlayHAlign = 0
          } else {
            overlayPinX = node.x * PX_TO_INCHES - txtNodeW / 2 + marginX - overlayW / 2
            overlayHAlign = 2
          }
        }
        const overlayPinY = pageHeightInches - node.y * PX_TO_INCHES
        shapes.push(`
      <Shape ID="${shapeId}" Type="Shape" NameU="TextOverlay.${shapeId}">
        <Cell N="PinX" V="${overlayPinX.toFixed(4)}" U="IN"/>
        <Cell N="PinY" V="${overlayPinY.toFixed(4)}" U="IN"/>
        <Cell N="Width" V="${overlayW.toFixed(4)}" U="IN"/>
        <Cell N="Height" V="${overlayH.toFixed(4)}" U="IN"/>
        <Cell N="LocPinX" V="${(overlayW / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="LocPinY" V="${(overlayH / 2).toFixed(4)}" U="IN" F="Height*0.5"/>
        <Cell N="FillPattern" V="0"/>
        <Cell N="LinePattern" V="0"/>
        <Cell N="TxtPinX" V="${(overlayW / 2).toFixed(4)}" U="IN" F="Width*0.5"/>
        <Cell N="TxtPinY" V="${(overlayH / 2).toFixed(4)}" U="IN" F="Height*0.5"/>
        <Cell N="TxtLocPinX" V="${(overlayW / 2).toFixed(4)}" U="IN" F="TxtWidth*0.5"/>
        <Cell N="TxtLocPinY" V="${(overlayH / 2).toFixed(4)}" U="IN" F="TxtHeight*0.5"/>
        <Cell N="TxtWidth" V="${overlayW.toFixed(4)}" U="IN" F="Width*1"/>
        <Cell N="TxtHeight" V="${overlayH.toFixed(4)}" U="IN" F="Height*1"/>
        <Cell N="VerticalAlign" V="${txtVAlign}"/>
        <Section N="Character">
          <Row IX="0">
            <Cell N="Font" V="Calibri"/>
            <Cell N="Color" V="${txtFontColor}"/>
            <Cell N="Size" V="${txtFontSize}" U="IN"/>
            <Cell N="Style" V="0"/>
          </Row>
        </Section>
        <Section N="Paragraph">
          <Row IX="0">
            <Cell N="HorzAlign" V="${overlayHAlign}"/>
          </Row>
        </Section>
        <Text><cp IX="0"/><pp IX="0"/>${escapeXml(node.label)}</Text>
      </Shape>`)
        shapeId++
      }
    })
  }

  // Build page1.xml.rels for image references
  let pageRelsXml = ''
  if (imageRels.length > 0) {
    const rels = imageRels
      .map(
        (r) =>
          `  <Relationship Id="${r.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${r.imageKey}.png"/>`
      )
      .join('\n')
    pageRelsXml = `<?xml version="1.0" encoding="utf-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${rels}\n</Relationships>`
  }

  const xml = `<?xml version="1.0" encoding="utf-8"?>
<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve">
  <PageSheet>
    <Cell N="PageWidth" V="${pageWidthInches.toFixed(4)}"/>
    <Cell N="PageHeight" V="${pageHeightInches.toFixed(4)}"/>
    <Cell N="PageScale" V="1"/>
    <Cell N="DrawingScale" V="1"/>
    <Cell N="DrawingSizeType" V="1"/>
    <Cell N="DrawingScaleType" V="0"/>
    <Cell N="InhibitSnap" V="0"/>
    <Cell N="ShdwType" V="0"/>
  </PageSheet>
  <Shapes>${shapes.join('')}
  </Shapes>${connectsSection}
</PageContents>`

  return { xml, pageWidthInches, pageHeightInches, pageRelsXml }
}
