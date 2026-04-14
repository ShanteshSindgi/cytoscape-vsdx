/**
 * Static XML templates for VSDX file parts.
 * A VSDX file is a ZIP archive following the Open Packaging Convention (OPC) with Visio-specific XML parts.
 */

/**
 * ZIP path: [Content_Types].xml
 * Manifest that maps file extensions and specific parts to their MIME types.
 * Visio reads this first to know how to interpret every file in the archive.
 * The exporter may append a PNG entry at runtime if the diagram has icons.
 */
export const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>
  <Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>
  <Override PartName="/visio/pages/page1.xml" ContentType="application/vnd.ms-visio.page+xml"/>
  <Override PartName="/visio/windows.xml" ContentType="application/vnd.ms-visio.windows+xml"/>
  <Override PartName="/visio/masters/masters.xml" ContentType="application/vnd.ms-visio.masters+xml"/>
  <Override PartName="/visio/masters/master1.xml" ContentType="application/vnd.ms-visio.master+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`

/**
 * ZIP path: _rels/.rels
 * Root relationship file — the entry point Visio reads on file open.
 * Maps rId1→document.xml, rId2→core.xml, rId3→app.xml.
 * These are resolved by Type URI, not explicitly referenced by other XML.
 */
export const ROOT_RELS_XML = `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document" Target="visio/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

/**
 * ZIP path: visio/document.xml
 * Main Visio document — defines default styles (line, fill, text),
 * registers the Calibri font, and sets document-level properties.
 * All shapes inherit from StyleSheet ID="0" unless overridden.
 */
export const DOCUMENT_XML = `<?xml version="1.0" encoding="utf-8"?>
<VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve">
  <DocumentProperties>
    <Creator>cytoscape-vsdx</Creator>
  </DocumentProperties>
  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0">
    <DynamicGridEnabled>1</DynamicGridEnabled>
  </DocumentSettings>
  <Colors/>
  <FaceNames>
    <FaceName ID="0" Name="Calibri" UnicodeRanges="-536870145 -1073732485 9 0" CharSets="536871423"/>
  </FaceNames>
  <StyleSheets>
    <StyleSheet ID="0" Name="No Style" NameU="No Style">
      <Cell N="LineWeight" V="0.01041666666666667"/>
      <Cell N="LineColor" V="#000000"/>
      <Cell N="LinePattern" V="1"/>
      <Cell N="FillForegnd" V="#FFFFFF"/>
      <Cell N="FillBkgnd" V="#000000"/>
      <Cell N="FillPattern" V="1"/>
      <Section N="Character">
        <Row IX="0">
          <Cell N="Font" V="0"/>
          <Cell N="Color" V="#000000"/>
          <Cell N="Size" V="0.1666666666666667"/>
        </Row>
      </Section>
    </StyleSheet>
  </StyleSheets>
</VisioDocument>`

/**
 * ZIP path: visio/_rels/document.xml.rels
 * Links document.xml to its child parts: pages, windows, and masters.
 */
export const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/pages" Target="pages/pages.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/windows" Target="windows.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/visio/2010/relationships/masters" Target="masters/masters.xml"/>
</Relationships>`

/**
 * ZIP path: visio/pages/pages.xml
 * Page index — declares a single page "Page-1".
 * The <Rel r:id="rId1"/> links to pages.xml.rels to find page1.xml.
 */
export const PAGES_XML = `<?xml version="1.0" encoding="utf-8"?>
<Pages xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve">
  <Page ID="0" Name="Page-1" NameU="Page-1">
    <Rel r:id="rId1"/>
  </Page>
</Pages>`

/**
 * ZIP path: visio/pages/_rels/pages.xml.rels
 * Links the page index entry (rId1) to the actual page content file (page1.xml).
 */
export const PAGES_RELS_XML = `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page1.xml"/>
</Relationships>`

/**
 * ZIP path: visio/masters/masters.xml
 * Master shape index — defines a single reusable "Rectangle" master.
 * Shapes in page1.xml use inline geometry rather than referencing this master,
 * but Visio requires non-empty masters to open the file without warnings.
 */
export const MASTERS_XML = `<?xml version="1.0" encoding="utf-8"?>
<Masters xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve">
  <Master ID="0" Name="Rectangle" NameU="Rectangle" UniqueID="{00000000-0000-0000-0000-000000000001}" IconSize="1" AlignName="2" MatchByName="1" IconUpdate="1">
    <Rel r:id="rId1"/>
  </Master>
</Masters>`

/**
 * ZIP path: visio/masters/_rels/masters.xml.rels
 * Links the masters index entry (rId1) to the master shape definition (master1.xml).
 */
export const MASTERS_RELS_XML = `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/master" Target="master1.xml"/>
</Relationships>`

/**
 * ZIP path: visio/masters/master1.xml
 * Rectangle master shape definition — a 1" × 0.75" rectangle with centered text.
 * Provides both V (static value) and F (formula) attributes so the shape
 * scales correctly if resized in Visio. Geometry draws a closed rectangle
 * via MoveTo + 4 LineTo rows.
 */
export const MASTER1_XML = `<?xml version="1.0" encoding="utf-8"?>
<MasterContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve">
  <Shapes>
    <Shape ID="1" Type="Shape" NameU="Rectangle">
      <Cell N="PinX" V="0.5" U="IN" F="Width*0.5"/>
      <Cell N="PinY" V="0.375" U="IN" F="Height*0.5"/>
      <Cell N="Width" V="1" U="IN"/>
      <Cell N="Height" V="0.75" U="IN"/>
      <Cell N="LocPinX" V="0.5" U="IN" F="Width*0.5"/>
      <Cell N="LocPinY" V="0.375" U="IN" F="Height*0.5"/>
      <Cell N="FillForegnd" V="#FFFFFF"/>
      <Cell N="FillBkgnd" V="#FFFFFF"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="LineColor" V="#000000"/>
      <Cell N="LineWeight" V="0.01" U="IN"/>
      <Cell N="LinePattern" V="1"/>
      <Cell N="TxtPinX" V="0.5" U="IN" F="Width*0.5"/>
      <Cell N="TxtPinY" V="0.375" U="IN" F="Height*0.5"/>
      <Cell N="TxtLocPinX" V="0.5" U="IN" F="TxtWidth*0.5"/>
      <Cell N="TxtLocPinY" V="0.375" U="IN" F="TxtHeight*0.5"/>
      <Cell N="TxtWidth" V="1" U="IN" F="Width*1"/>
      <Cell N="TxtHeight" V="0.75" U="IN" F="Height*1"/>
      <Cell N="VerticalAlign" V="1"/>
      <Section N="Character">
        <Row IX="0">
          <Cell N="Font" V="0"/>
          <Cell N="Color" V="#000000"/>
          <Cell N="Size" V="0.1667" U="IN"/>
        </Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0">
          <Cell N="HorzAlign" V="1"/>
        </Row>
      </Section>
      <Section N="Geometry" IX="0">
        <Cell N="NoFill" V="0"/>
        <Cell N="NoLine" V="0"/>
        <Row T="MoveTo" IX="1">
          <Cell N="X" V="0" F="Width*0"/>
          <Cell N="Y" V="0" F="Height*0"/>
        </Row>
        <Row T="LineTo" IX="2">
          <Cell N="X" V="1" F="Width*1"/>
          <Cell N="Y" V="0" F="Height*0"/>
        </Row>
        <Row T="LineTo" IX="3">
          <Cell N="X" V="1" F="Width*1"/>
          <Cell N="Y" V="0.75" F="Height*1"/>
        </Row>
        <Row T="LineTo" IX="4">
          <Cell N="X" V="0" F="Width*0"/>
          <Cell N="Y" V="0.75" F="Height*1"/>
        </Row>
        <Row T="LineTo" IX="5">
          <Cell N="X" V="0" F="Width*0"/>
          <Cell N="Y" V="0" F="Height*0"/>
        </Row>
      </Section>
      <Text/>
    </Shape>
  </Shapes>
</MasterContents>`

/**
 * ZIP path: visio/windows.xml
 * Builds viewport/window settings for Visio desktop. Centers the view on the
 * middle of the page and hides grid, guides, connection points, and page breaks
 * for a clean initial view.
 *
 * @param pageWidthInches Total page width in inches (from buildPageXml)
 * @param pageHeightInches Total page height in inches (from buildPageXml)
 * @returns Complete windows.xml content
 */
export function buildWindowsXml(pageWidthInches: number, pageHeightInches: number): string {
  return `<?xml version="1.0" encoding="utf-8"?>
<Windows xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ClientWidth="1000" ClientHeight="700">
  <Window ID="0" WindowType="Drawing" WindowState="1073742340" WindowLeft="0" WindowTop="0" WindowWidth="1000" WindowHeight="700"
    ContainerType="Page" Page="0" ViewScale="1" ViewCenterX="${(pageWidthInches / 2).toFixed(4)}" ViewCenterY="${(
    pageHeightInches / 2
  ).toFixed(4)}">
    <ShowGrid>0</ShowGrid>
    <ShowGuides>0</ShowGuides>
    <ShowConnectionPoints>0</ShowConnectionPoints>
    <ShowPageBreaks>0</ShowPageBreaks>
  </Window>
</Windows>`
}

/**
 * ZIP path: docProps/app.xml
 * Application-level metadata. Declares the file was created by "Microsoft Visio"
 * v15 and company. The Application and AppVersion fields ensure
 * compatibility with Visio desktop, Lucidchart, and draw.io.
 *
 * @returns Complete app.xml content
 */
export function buildAppXml(): string {
  return `<?xml version="1.0" encoding="utf-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template/>
  <Application>Microsoft Visio</Application>
  <ScaleCrop>false</ScaleCrop>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>15.0000</AppVersion>
</Properties>`
}

/**
 * ZIP path: docProps/core.xml
 * Dublin Core metadata — sets creator to "cytoscape-vsdx" and
 * stamps the current date/time as both created and modified timestamps.
 *
 * @returns Complete core.xml content
 */
export function buildCoreXml(): string {
  const now = new Date().toISOString()
  return `<?xml version="1.0" encoding="utf-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title/>
  <dc:subject/>
  <dc:creator>cytoscape-vsdx</dc:creator>
  <dc:description/>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
  <cp:category/>
  <dc:language>en-US</dc:language>
</cp:coreProperties>`
}
