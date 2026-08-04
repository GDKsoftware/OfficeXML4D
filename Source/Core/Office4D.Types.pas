unit Office4D.Types;

interface

const
  // Content Types
  ContentTypeRelationships = 'application/vnd.openxmlformats-package.relationships+xml';
  ContentTypeWordDocument = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
  ContentTypeExcelWorkbook = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';

  // Relationship Types
  RelTypeOfficeDocument = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RelTypeWorksheet = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  RelTypeSharedStrings = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
  RelTypeStyles = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  RelTypeCoreProperties = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

  // XML Namespaces
  NsPackageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
  NsOfficeDocumentRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  NsContentTypes = 'http://schemas.openxmlformats.org/package/2006/content-types';
  NsWordprocessingML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  NsSpreadsheetML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  NsDublinCore = 'http://purl.org/dc/elements/1.1/';
  NsDcTerms = 'http://purl.org/dc/terms/';
  NsCoreProperties = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

  // Every part this library writes opens with this declaration
  XmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  // Standard part paths
  PartContentTypes = '[Content_Types].xml';
  PartRootRels = '_rels/.rels';
  PartCoreProperties = 'docProps/core.xml';
  PartAppProperties = 'docProps/app.xml';

implementation

end.
