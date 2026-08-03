unit Office4D.Word.Writer;

interface

uses
  System.Generics.Collections,
  System.Zip,
  Office4D.Word,
  Office4D.Word.Model;

type
  /// <summary>
  /// Turns a TWordDocumentContent into the parts of a .docx and adds them to a
  /// zip. Generating the XML and deciding which parts exist belong together:
  /// the relationship ids in document.xml and document.xml.rels are counted out
  /// on both sides and have to agree.
  /// </summary>
  TWordDocumentWriter = class
  private
    FContent: TWordDocumentContent;

    function GenerateContentTypesXml: string;
    function GenerateRootRelsXml: string;
    function GenerateDocumentXml: string;
    function GenerateDocumentRelsXml: string;
    function GenerateHeaderXml: string;
    function GenerateFooterXml: string;
    function GenerateNumberingXml: string;
    function GenerateBorderXml(const ElementName: string; const Border: TTableBorder): string;
    function GenerateDrawingXml(const ImageRun: TWordRun; ImageIndex: Integer; RelIdOffset: Integer): string;
    function CollectHyperlinks: TList<string>;
    function GetHyperlinkId(const Url: string; const Hyperlinks: TList<string>): Integer;
    function CollectImages: TList<TWordRun>;
    function GetImageContentType(const Extension: string): string;
    function HasListParagraphs: Boolean;
    function HasHeader: Boolean;
    function HasFooter: Boolean;
  public
    constructor Create(const Content: TWordDocumentContent);

    procedure WriteParts(const Zip: TZipFile);
  end;

implementation

uses
  System.SysUtils,
  Office4D.Types,
  Office4D.Xml;

const
  WordprocessingNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  DrawingNs = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  WpDrawingNs = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  PictureNs = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

  PartDocumentRels = 'word/_rels/document.xml.rels';

{ TWordDocumentWriter }

constructor TWordDocumentWriter.Create(const Content: TWordDocumentContent);
begin
  inherited Create;
  FContent := Content;
end;

function TWordDocumentWriter.CollectHyperlinks: TList<string>;
begin
  Result := TList<string>.Create;
  for var Para in FContent.Paragraphs do
  begin
    for var RunIndex := 0 to Para.RunCount - 1 do
    begin
      var Run := Para.Runs[RunIndex];
      if (Run.Hyperlink <> '') and (not Result.Contains(Run.Hyperlink)) then
        Result.Add(Run.Hyperlink);
    end;
  end;
end;

function TWordDocumentWriter.GetHyperlinkId(const Url: string; const Hyperlinks: TList<string>): Integer;
begin
  Result := Hyperlinks.IndexOf(Url) + 1;
end;

function TWordDocumentWriter.HasListParagraphs: Boolean;
begin
  for var Para in FContent.Paragraphs do
    if Para.ListStyle <> TListStyle.None then
      Exit(True);
  Result := False;
end;

function TWordDocumentWriter.GenerateNumberingXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<w:numbering xmlns:w="' + WordprocessingNs + '">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0">' +
    '<w:start w:val="1"/>' +
    '<w:numFmt w:val="bullet"/>' +
    '<w:lvlText w:val=""/>' +
    '<w:lvlJc w:val="left"/>' +
    '</w:lvl>' +
    '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1">' +
    '<w:lvl w:ilvl="0">' +
    '<w:start w:val="1"/>' +
    '<w:numFmt w:val="decimal"/>' +
    '<w:lvlText w:val="%1."/>' +
    '<w:lvlJc w:val="left"/>' +
    '</w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>';
end;

function TWordDocumentWriter.GenerateContentTypesXml: string;
begin
  var NumberingOverride := '';
  if HasListParagraphs then
    NumberingOverride := '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';

  var HeaderOverride := '';
  if HasHeader then
    HeaderOverride := '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';

  var FooterOverride := '';
  if HasFooter then
    FooterOverride := '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';

  var ImageDefaults := '';
  var Images := CollectImages;
  try
    var Extensions := TList<string>.Create;
    try
      for var ImgRun in Images do
      begin
        var Ext := LowerCase(ImgRun.GetImage.Extension);
        if not Extensions.Contains(Ext) then
        begin
          Extensions.Add(Ext);
          ImageDefaults := ImageDefaults + '<Default Extension="' + Ext + '" ContentType="' + GetImageContentType(Ext) + '"/>';
        end;
      end;
    finally
      Extensions.Free;
    end;
  finally
    Images.Free;
  end;

  Result :=
    XmlDeclaration + sLineBreak +
    '<Types xmlns="' + NsContentTypes + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    ImageDefaults +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    NumberingOverride +
    HeaderOverride +
    FooterOverride +
    '</Types>';
end;

function TWordDocumentWriter.GenerateRootRelsXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<Relationships xmlns="' + NsPackageRelationships + '">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
end;

function TWordDocumentWriter.GenerateDocumentRelsXml: string;
begin
  var Hyperlinks := CollectHyperlinks;
  var Images := CollectImages;
  try
    var RelsContent := '';
    var NextId := 1;

    if HasListParagraphs then
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" ' +
        'Target="numbering.xml"/>';
      Inc(NextId);
    end;

    if HasHeader then
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ' +
        'Target="header1.xml"/>';
      Inc(NextId);
    end;

    if HasFooter then
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" ' +
        'Target="footer1.xml"/>';
      Inc(NextId);
    end;

    for var I := 0 to Hyperlinks.Count - 1 do
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId + I) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
        'Target="' + TXml.Escape(Hyperlinks[I]) + '" TargetMode="External"/>';
    end;
    NextId := NextId + Hyperlinks.Count;

    for var I := 0 to Images.Count - 1 do
    begin
      var ImgRun := Images[I];
      var Ext := LowerCase(ImgRun.GetImage.Extension);
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId + I) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="media/image' + IntToStr(I + 1) + '.' + Ext + '"/>';
    end;

    Result :=
      XmlDeclaration + sLineBreak +
      '<Relationships xmlns="' + NsPackageRelationships + '">' +
      RelsContent +
      '</Relationships>';
  finally
    Hyperlinks.Free;
    Images.Free;
  end;
end;

function TWordDocumentWriter.GenerateDocumentXml: string;
begin
  var Hyperlinks := CollectHyperlinks;
  var Images := CollectImages;
  try
    var Body := '';
    var HyperlinkIdOffset := 0;
    if HasListParagraphs then
      Inc(HyperlinkIdOffset);
    if HasHeader then
      Inc(HyperlinkIdOffset);
    if HasFooter then
      Inc(HyperlinkIdOffset);

    var ImageIdOffset := HyperlinkIdOffset + Hyperlinks.Count;
    var ImageIndex := 0;

    for var I := 0 to FContent.Paragraphs.Count - 1 do
    begin
      var Para := FContent.Paragraphs[I];
      Body := Body + '<w:p>';

      const HasSpacing = (Para.LineSpacing.Line > 0) or (Para.LineSpacing.Before > 0) or (Para.LineSpacing.After > 0);
      const HasIndent = (Para.Indent.Left > 0) or (Para.Indent.Right > 0) or (Para.Indent.FirstLine <> 0);
      const NeedPPr = (Para.ListStyle <> TListStyle.None) or (Para.Alignment <> TParagraphAlignment.Left) or HasSpacing or HasIndent;
      if NeedPPr then
      begin
        Body := Body + '<w:pPr>';
        if Para.ListStyle <> TListStyle.None then
        begin
          var NumId := '1';
          if Para.ListStyle = TListStyle.Numbered then
            NumId := '2';
          Body := Body + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + NumId + '"/></w:numPr>';
        end;
        if HasSpacing then
        begin
          var SpacingAttrs := '';
          if Para.LineSpacing.Before > 0 then
            SpacingAttrs := SpacingAttrs + ' w:before="' + IntToStr(Para.LineSpacing.Before) + '"';
          if Para.LineSpacing.After > 0 then
            SpacingAttrs := SpacingAttrs + ' w:after="' + IntToStr(Para.LineSpacing.After) + '"';
          if Para.LineSpacing.Line > 0 then
          begin
            SpacingAttrs := SpacingAttrs + ' w:line="' + IntToStr(Para.LineSpacing.Line) + '"';
            var RuleValue := 'auto';
            case Para.LineSpacing.Rule of
              TLineSpacingRule.Exact: RuleValue := 'exact';
              TLineSpacingRule.AtLeast: RuleValue := 'atLeast';
            end;
            SpacingAttrs := SpacingAttrs + ' w:lineRule="' + RuleValue + '"';
          end;
          Body := Body + '<w:spacing' + SpacingAttrs + '/>';
        end;
        if HasIndent then
        begin
          var IndentAttrs := '';
          if Para.Indent.Left > 0 then
            IndentAttrs := IndentAttrs + ' w:left="' + IntToStr(Para.Indent.Left) + '"';
          if Para.Indent.Right > 0 then
            IndentAttrs := IndentAttrs + ' w:right="' + IntToStr(Para.Indent.Right) + '"';
          if Para.Indent.FirstLine > 0 then
            IndentAttrs := IndentAttrs + ' w:firstLine="' + IntToStr(Para.Indent.FirstLine) + '"'
          else if Para.Indent.FirstLine < 0 then
            IndentAttrs := IndentAttrs + ' w:hanging="' + IntToStr(-Para.Indent.FirstLine) + '"';
          Body := Body + '<w:ind' + IndentAttrs + '/>';
        end;
        if Para.Alignment <> TParagraphAlignment.Left then
        begin
          var AlignValue := 'left';
          case Para.Alignment of
            TParagraphAlignment.Center: AlignValue := 'center';
            TParagraphAlignment.Right: AlignValue := 'right';
            TParagraphAlignment.Justify: AlignValue := 'both';
          end;
          Body := Body + '<w:jc w:val="' + AlignValue + '"/>';
        end;
        Body := Body + '</w:pPr>';
      end;

      for var J := 0 to Para.RunCount - 1 do
      begin
        var Run := Para.Runs[J];
        var WordRun := Run as TWordRun;
        var RunText := Run.Text;
        var HasHyperlink := Run.Hyperlink <> '';

        if WordRun.HasImage then
        begin
          Inc(ImageIndex);
          Body := Body + '<w:r>' + GenerateDrawingXml(WordRun, ImageIndex, ImageIdOffset) + '</w:r>';
        end
        else if RunText = sLineBreak then
          Body := Body + '<w:r><w:br/></w:r>'
        else if RunText = #9 then
          Body := Body + '<w:r><w:tab/></w:r>'
        else if RunText = #12 then
          Body := Body + '<w:r><w:br w:type="page"/></w:r>'
        else
        begin
          if HasHyperlink then
            Body := Body + '<w:hyperlink r:id="rId' + IntToStr(GetHyperlinkId(Run.Hyperlink, Hyperlinks) + HyperlinkIdOffset) + '">';

          Body := Body + '<w:r>';
          const NeedRPr = Run.Bold or Run.Italic or Run.Underline or
                         (Run.FontName <> '') or (Run.FontSize > 0) or (Run.FontColor <> '');
          if NeedRPr then
          begin
            Body := Body + '<w:rPr>';
            if Run.FontName <> '' then
              Body := Body + '<w:rFonts w:ascii="' + TXml.Escape(Run.FontName) + '" w:hAnsi="' + TXml.Escape(Run.FontName) + '"/>';
            if Run.Bold then
              Body := Body + '<w:b/>';
            if Run.Italic then
              Body := Body + '<w:i/>';
            if Run.Underline then
              Body := Body + '<w:u w:val="single"/>';
            if Run.FontSize > 0 then
              Body := Body + '<w:sz w:val="' + IntToStr(Run.FontSize) + '"/>';
            if Run.FontColor <> '' then
              Body := Body + '<w:color w:val="' + TXml.Escape(Run.FontColor) + '"/>';
            Body := Body + '</w:rPr>';
          end;
          Body := Body + '<w:t xml:space="preserve">' + TXml.Escape(RunText) + '</w:t></w:r>';

          if HasHyperlink then
            Body := Body + '</w:hyperlink>';
        end;
      end;

      Body := Body + '</w:p>';
    end;

    for var I := 0 to FContent.Tables.Count - 1 do
    begin
      var Table := FContent.Tables[I];
      var WordTable := Table as TWordTable;
      Body := Body + '<w:tbl>';

      Body := Body + '<w:tblPr>';
      Body := Body + '<w:tblW w:w="0" w:type="auto"/>';

      var Borders := WordTable.GetBorders;
      const HasBorders = (Borders.Top.Style <> TBorderStyle.None) or
                         (Borders.Bottom.Style <> TBorderStyle.None) or
                         (Borders.Left.Style <> TBorderStyle.None) or
                         (Borders.Right.Style <> TBorderStyle.None) or
                         (Borders.InsideH.Style <> TBorderStyle.None) or
                         (Borders.InsideV.Style <> TBorderStyle.None);
      if HasBorders then
      begin
        Body := Body + '<w:tblBorders>';
        if Borders.Top.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:top', Borders.Top);
        if Borders.Left.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:left', Borders.Left);
        if Borders.Bottom.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:bottom', Borders.Bottom);
        if Borders.Right.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:right', Borders.Right);
        if Borders.InsideH.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:insideH', Borders.InsideH);
        if Borders.InsideV.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:insideV', Borders.InsideV);
        Body := Body + '</w:tblBorders>';
      end;
      Body := Body + '</w:tblPr>';

      var ColumnWidths := WordTable.GetColumnWidths;
      if Length(ColumnWidths) > 0 then
      begin
        Body := Body + '<w:tblGrid>';
        for var ColW in ColumnWidths do
          Body := Body + '<w:gridCol w:w="' + IntToStr(ColW) + '"/>';
        Body := Body + '</w:tblGrid>';
      end;

      for var R := 0 to Table.RowCount - 1 do
      begin
        Body := Body + '<w:tr>';
        for var C := 0 to Table.ColCount - 1 do
        begin
          var Cell := Table.Cells[R, C];
          var WordCell := Cell as TWordTableCell;
          Body := Body + '<w:tc>';

          const HasCellPr = (WordCell.GetShading <> '') or (WordCell.GetWidth > 0);
          if HasCellPr then
          begin
            Body := Body + '<w:tcPr>';
            if WordCell.GetWidth > 0 then
              Body := Body + '<w:tcW w:w="' + IntToStr(WordCell.GetWidth) + '" w:type="dxa"/>';
            if WordCell.GetShading <> '' then
              Body := Body + '<w:shd w:val="clear" w:fill="' + WordCell.GetShading + '"/>';
            Body := Body + '</w:tcPr>';
          end;

          Body := Body + '<w:p><w:r><w:t>' + TXml.Escape(Cell.Text) + '</w:t></w:r></w:p>';
          Body := Body + '</w:tc>';
        end;
        Body := Body + '</w:tr>';
      end;

      Body := Body + '</w:tbl>';
    end;

    var SectPr := '';
    var NeedSectPr := (FContent.PageOrientation = TPageOrientation.Landscape) or
       (FContent.PageMargins.Top > 0) or (FContent.PageMargins.Bottom > 0) or
       (FContent.PageMargins.Left > 0) or (FContent.PageMargins.Right > 0) or
       HasHeader or HasFooter;

    if NeedSectPr then
    begin
      SectPr := '<w:sectPr>';

      var HeaderFooterRelId := 1;
      if HasListParagraphs then
        Inc(HeaderFooterRelId);

      if HasHeader then
      begin
        SectPr := SectPr + '<w:headerReference w:type="default" r:id="rId' + IntToStr(HeaderFooterRelId) + '"/>';
        Inc(HeaderFooterRelId);
      end;

      if HasFooter then
        SectPr := SectPr + '<w:footerReference w:type="default" r:id="rId' + IntToStr(HeaderFooterRelId) + '"/>';

      if FContent.PageOrientation = TPageOrientation.Landscape then
        SectPr := SectPr + '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
      else if NeedSectPr then
        SectPr := SectPr + '<w:pgSz w:w="12240" w:h="15840"/>';

      if (FContent.PageMargins.Top > 0) or (FContent.PageMargins.Bottom > 0) or
         (FContent.PageMargins.Left > 0) or (FContent.PageMargins.Right > 0) then
      begin
        SectPr := SectPr + '<w:pgMar w:top="' + IntToStr(FContent.PageMargins.Top) + '" ' +
          'w:bottom="' + IntToStr(FContent.PageMargins.Bottom) + '" ' +
          'w:left="' + IntToStr(FContent.PageMargins.Left) + '" ' +
          'w:right="' + IntToStr(FContent.PageMargins.Right) + '"/>';
      end;
      SectPr := SectPr + '</w:sectPr>';
    end;

    Result :=
      XmlDeclaration + sLineBreak +
      '<w:document xmlns:w="' + WordprocessingNs + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<w:body>' + Body + SectPr + '</w:body>' +
      '</w:document>';
  finally
    Hyperlinks.Free;
    Images.Free;
  end;
end;

procedure TWordDocumentWriter.WriteParts(const Zip: TZipFile);
begin
  var ContentTypesXml := GenerateContentTypesXml;
  var RelsXml := GenerateRootRelsXml;
  var DocumentXml := GenerateDocumentXml;
  var DocumentRelsXml := GenerateDocumentRelsXml;

  Zip.Add(TEncoding.UTF8.GetBytes(ContentTypesXml), '[Content_Types].xml');
  Zip.Add(TEncoding.UTF8.GetBytes(RelsXml), PartRootRels);
  Zip.Add(TEncoding.UTF8.GetBytes(DocumentXml), 'word/document.xml');
  Zip.Add(TEncoding.UTF8.GetBytes(DocumentRelsXml), PartDocumentRels);

  if HasListParagraphs then
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateNumberingXml), 'word/numbering.xml');

  if HasHeader then
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateHeaderXml), 'word/header1.xml');

  if HasFooter then
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateFooterXml), 'word/footer1.xml');

  var Images := CollectImages;
  try
    for var Idx := 0 to Images.Count - 1 do
    begin
      var Img := Images[Idx].GetImage;
      var MediaPath := 'word/media/image' + IntToStr(Idx + 1) + '.' + LowerCase(Img.Extension);
      Zip.Add(Img.Data, MediaPath);
    end;
  finally
    Images.Free;
  end;
end;

function TWordDocumentWriter.HasHeader: Boolean;
begin
  Result := FContent.Header.Text <> '';
end;

function TWordDocumentWriter.HasFooter: Boolean;
begin
  Result := FContent.Footer.Text <> '';
end;

function TWordDocumentWriter.GenerateBorderXml(const ElementName: string; const Border: TTableBorder): string;
begin
  var StyleValue := 'single';
  case Border.Style of
    TBorderStyle.Single: StyleValue := 'single';
    TBorderStyle.Double: StyleValue := 'double';
    TBorderStyle.Dashed: StyleValue := 'dashed';
    TBorderStyle.Dotted: StyleValue := 'dotted';
    TBorderStyle.Thick: StyleValue := 'thick';
  end;
  Result := '<' + ElementName + ' w:val="' + StyleValue + '" w:sz="' + IntToStr(Border.Width) + '"';
  if Border.Color <> '' then
    Result := Result + ' w:color="' + Border.Color + '"';
  Result := Result + '/>';
end;

function TWordDocumentWriter.CollectImages: TList<TWordRun>;
begin
  Result := TList<TWordRun>.Create;
  for var Para in FContent.Paragraphs do
  begin
    for var RunIndex := 0 to Para.RunCount - 1 do
    begin
      var Run := Para.Runs[RunIndex];
      var WordRun := Run as TWordRun;
      if WordRun.HasImage then
        Result.Add(WordRun);
    end;
  end;
end;

function TWordDocumentWriter.GetImageContentType(const Extension: string): string;
begin
  var Ext := LowerCase(Extension);
  if (Ext = 'jpg') or (Ext = 'jpeg') then
    Result := 'image/jpeg'
  else if Ext = 'png' then
    Result := 'image/png'
  else if Ext = 'gif' then
    Result := 'image/gif'
  else if Ext = 'bmp' then
    Result := 'image/bmp'
  else
    Result := 'image/png';
end;

function TWordDocumentWriter.GenerateDrawingXml(const ImageRun: TWordRun; ImageIndex: Integer; RelIdOffset: Integer): string;
begin
  var Img := ImageRun.GetImage;
  var RelId := 'rId' + IntToStr(ImageIndex + RelIdOffset);
  var ImgName := 'Image ' + IntToStr(ImageIndex);
  var CxStr := IntToStr(Img.Width);
  var CyStr := IntToStr(Img.Height);

  Result :=
    '<w:drawing>' +
      '<wp:inline xmlns:wp="' + WpDrawingNs + '" distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="' + CxStr + '" cy="' + CyStr + '"/>' +
        '<wp:docPr id="' + IntToStr(ImageIndex) + '" name="' + ImgName + '"/>' +
        '<a:graphic xmlns:a="' + DrawingNs + '">' +
          '<a:graphicData uri="' + PictureNs + '">' +
            '<pic:pic xmlns:pic="' + PictureNs + '">' +
              '<pic:nvPicPr>' +
                '<pic:cNvPr id="' + IntToStr(ImageIndex) + '" name="' + ImgName + '"/>' +
                '<pic:cNvPicPr/>' +
              '</pic:nvPicPr>' +
              '<pic:blipFill>' +
                '<a:blip xmlns:r="' + NsOfficeDocumentRelationships + '" r:embed="' + RelId + '"/>' +
                '<a:stretch><a:fillRect/></a:stretch>' +
              '</pic:blipFill>' +
              '<pic:spPr>' +
                '<a:xfrm>' +
                  '<a:off x="0" y="0"/>' +
                  '<a:ext cx="' + CxStr + '" cy="' + CyStr + '"/>' +
                '</a:xfrm>' +
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
              '</pic:spPr>' +
            '</pic:pic>' +
          '</a:graphicData>' +
        '</a:graphic>' +
      '</wp:inline>' +
    '</w:drawing>';
end;

function TWordDocumentWriter.GenerateHeaderXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<w:hdr xmlns:w="' + WordprocessingNs + '">' +
    '<w:p><w:r><w:t>' + TXml.Escape(FContent.Header.Text) + '</w:t></w:r></w:p>' +
    '</w:hdr>';
end;

function TWordDocumentWriter.GenerateFooterXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<w:ftr xmlns:w="' + WordprocessingNs + '">' +
    '<w:p><w:r><w:t>' + TXml.Escape(FContent.Footer.Text) + '</w:t></w:r></w:p>' +
    '</w:ftr>';
end;

end.
