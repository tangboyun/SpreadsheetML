{-# LANGUAGE BangPatterns      #-}
module Text.XML.SpreadsheetML.Writer
       (
         showSpreadsheet
       , ToElement(..)
       )
       where

import           Data.Char
import qualified Text.XML.Light as L
import qualified Text.XML.Light.Output as O
import qualified Text.XML.Light.Types as LT
import qualified Text.XML.SpreadsheetML.Internal as I
import qualified Text.XML.SpreadsheetML.Types as T

import           Control.Applicative ( (<$>) )
import           Data.Maybe ( catMaybes, maybeToList )
import           Data.Colour.SRGB

--------------------------------------------------------------------------
-- | Convert a workbook to a string.  Write this string to a ".xls" file
-- and Excel will know how to open it.
showSpreadsheet :: T.Workbook -> String
showSpreadsheet wb = "<?xml version='1.0' ?>\n" ++
                     "<?mso-application progid=\"Excel.Sheet\"?>\n" ++
                     O.showElement (toElement wb)
                     
---------------------------------------------------------------------------
-- | Namespaces
namespace, oNamespace, xNamespace, ssNamespace, htmlNamespace :: LT.QName
namespace     = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:spreadsheet" }
oNamespace    = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:office"
                             , L.qPrefix = Just "o" }
xNamespace    = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:excel"
                             , L.qPrefix = Just "x" }
ssNamespace   = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:spreadsheet"
                             , L.qPrefix = Just "ss" }
htmlNamespace = L.blank_name { L.qURI    = Just "http://www.w3.org/TR/REC-html40"
                             , L.qPrefix = Just "html" }

--------------------------------------------------------------------------
-- | Empty Elements
emptyWorkbook :: LT.Element
emptyWorkbook = L.blank_element
  { L.elName    = workbookName
  , L.elAttribs = [xmlns, xmlns_o, xmlns_x, xmlns_ss, xmlns_html] }
  where
  workbookName = namespace { L.qName = "Workbook" }
  xmlns      = mkAttr "xmlns"      "urn:schemas-microsoft-com:office:spreadsheet"
  xmlns_o    = mkAttr "xmlns:o"    "urn:schemas-microsoft-com:office:office"
  xmlns_x    = mkAttr "xmlns:x"    "urn:schemas-microsoft-com:office:excel"
  xmlns_ss   = mkAttr "xmlns:ss"   "urn:schemas-microsoft-com:office:spreadsheet"
  xmlns_html = mkAttr "xmlns:html" "http://www.w3.org/TR/REC-html40"
  mkAttr !k !v = LT.Attr L.blank_name { L.qName = k } v

emptyDocumentProperties :: LT.Element
emptyDocumentProperties = L.blank_element { L.elName = documentPropertiesName }
  where
  documentPropertiesName = oNamespace { L.qName = "DocumentProperties" }

emptyWorksheet :: T.Name -> LT.Element
emptyWorksheet (T.Name !n) = L.blank_element { L.elName    = worksheetName
                                             , L.elAttribs = [LT.Attr worksheetNameAttrName n] }
  where
  worksheetName = namespace { L.qName   = "Worksheet" }
  worksheetNameAttrName = ssNamespace { L.qName   = "Name" }

emptyTable :: LT.Element
emptyTable = L.blank_element { L.elName = tableName }
  where
  tableName = namespace { L.qName = "Table" }

emptyRow :: LT.Element
emptyRow = L.blank_element { L.elName = rowName }
  where
  rowName = namespace { L.qName = "Row" }

emptyColumn :: LT.Element
emptyColumn = L.blank_element { L.elName = columnName }
  where
  columnName = namespace { L.qName = "Column" }

emptyCell :: LT.Element
emptyCell = L.blank_element { L.elName = cellName }
  where
  cellName = namespace { L.qName = "Cell" }

mkData :: I.ExcelValue -> LT.Element
mkData !v =
  case v of
    I.ExcelValue _ ->
      L.blank_element { L.elName     = richName
                      , L.elContent  = [ LT.Text (mkCData v) ]
                      , L.elAttribs  = [  mkAttr v 
                                       , LT.Attr (LT.blank_name { LT.qName = "xmlns" })
                                                 "http://www.w3.org/TR/REC-html40"]
                      }
    _ ->
      L.blank_element { L.elName     = dataName
                      , L.elContent  = [ LT.Text (mkCData v) ]
                      , L.elAttribs  = [ mkAttr v ] }
  where
  dataName   = namespace { L.qName = "Data" }
  richNamespace = L.blank_name { L.qPrefix = Just "ss" }
  richName = richNamespace { L.qName = "Data" }
  typeName !s = ssNamespace { L.qName = s }
  typeAttr   = LT.Attr (typeName "Type")
  mkAttr (I.Number _)      = typeAttr "Number"
  mkAttr (I.Boolean _)     = typeAttr "Boolean"
  mkAttr (I.StringType _)  = typeAttr "String"
  mkAttr (I.ExcelValue _)  = typeAttr "String"
  mkCData (I.Number !d)     = L.blank_cdata { LT.cdData = show d }
  mkCData (I.Boolean !b)    = L.blank_cdata { LT.cdData = showBoolean b } 
  mkCData (I.StringType !s) = L.blank_cdata { LT.cdData = s }
  mkCData (I.ExcelValue !richText) = L.blank_cdata { LT.cdData = I.fromRich richText, LT.cdVerbatim = LT.CDataRaw } 
              
-------------------------------------------------------------------------
-- | XML Conversion Class
class ToElement a where
  toElement :: a -> LT.Element

-------------------------------------------------------------------------
-- | Instances
instance ToElement T.Workbook where
  toElement wb = emptyWorkbook
    { L.elContent = mbook ++
                    maybeToList (LT.Elem . toElement <$> T.worksheetStyles wb) ++
                    map (LT.Elem . toElement) (T.workbookWorksheets wb) }
    where
    mbook = maybeToList (LT.Elem . toElement <$> T.workbookDocumentProperties wb)

instance ToElement T.DocumentProperties where
  toElement dp = emptyDocumentProperties
    { L.elContent = map LT.Elem $ catMaybes
      [ toE T.documentPropertiesTitle       "Title"       id
      , toE T.documentPropertiesSubject     "Subject"     id
      , toE T.documentPropertiesKeywords    "Keywords"    id
      , toE T.documentPropertiesDescription "Description" id
      , toE T.documentPropertiesRevision    "Revision"    show
      , toE T.documentPropertiesAppName     "AppName"     id
      , toE T.documentPropertiesCreated     "Created"     id
      ]
    }
    where
    toE :: (T.DocumentProperties -> Maybe a) -> String -> (a -> String) -> Maybe L.Element
    toE fieldOf !name toString = mkCData <$> fieldOf dp
      where
      mkCData !cdata = L.blank_element
        { L.elName    = oNamespace { L.qName = name }
        , L.elContent = [LT.Text (L.blank_cdata { L.cdData = toString cdata })] }

  
instance ToElement T.Worksheet where
  toElement ws = (emptyWorksheet (T.worksheetName ws))
    { L.elContent = maybeToList (LT.Elem . toElement <$> (T.worksheetTable ws))  }

instance ToElement T.Table where
  toElement t = emptyTable
    { L.elContent = map LT.Elem $
                    map toElement (T.tableColumns t) ++
                    map toElement (T.tableRows t)
    , L.elAttribs = catMaybes
                    [ toA T.tableDefaultColumnWidth  "DefaultColumnWidth"  show
                    , toA T.tableDefaultRowHeight    "DefaultRowHeight"    show
                    , toA T.tableExpandedColumnCount "ExpandedColumnCount" show
                    , toA T.tableExpandedRowCount    "ExpandedRowCount"    show
                    , toA T.tableLeftCell            "LeftCell"            show
                    , toA T.tableFullColumns         "FullColumns"         showBoolean
                    , toA T.tableFullRows            "FullRows"            showBoolean
                    , toA T.tableStyleID             "StyleID"             id
                    ] }
    where
    toA :: (T.Table -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf !name toString = mkAttr <$> fieldOf t
      where
      mkAttr !value = LT.Attr ssNamespace { L.qName = name } (toString value)

instance ToElement T.Row where
  toElement r = emptyRow
    { L.elContent = map (LT.Elem . toElement) (T.rowCells r)
    , L.elAttribs = catMaybes
      [ toA T.rowCaption       "Caption"       showCaption
      , toA T.rowAutoFitHeight "AutoFitHeight" showAutoFitHeight
      , toA T.rowHeight        "Height"        show
      , toA T.rowHidden        "Hidden"        showHidden
      , toA T.rowIndex         "Index"         show
      , toA T.rowSpan          "Span"          show
      , toA T.rowStyleID       "StyleID"       id
      ] }

    where
    showAutoFitHeight T.AutoFitHeight      = "1"
    showAutoFitHeight T.DoNotAutoFitHeight = "0"
    toA :: (T.Row -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf !name toString = mkAttr <$> fieldOf r
      where
      mkAttr !value = LT.Attr ssNamespace { L.qName = name } (toString value)

showBoolean :: Bool -> String
showBoolean True  = "1"
showBoolean False = "0"

showCaption :: T.Caption -> String
showCaption (T.Caption !s) = s

showHidden :: T.Hidden -> String
showHidden T.Hidden = "1"
showHidden T.Shown  = "0"

instance ToElement T.Column where
  toElement !c = emptyColumn
    { L.elAttribs = catMaybes
      [ toA T.columnCaption      "Caption"      showCaption
      , toA T.columnAutoFitWidth "AutoFitWidth" showAutoFitWidth
      , toA T.columnHidden       "Hidden"       showHidden
      , toA T.columnIndex        "Index"        show
      , toA T.columnSpan         "Span"         show
      , toA T.columnWidth        "Width"        show
      , toA T.columnStyleID      "StyleID"      id
      ] }
    where
    showAutoFitWidth T.AutoFitWidth      = "1"
    showAutoFitWidth T.DoNotAutoFitWidth = "0"
    toA :: (T.Column -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf !name toString = mkAttr <$> fieldOf c
      where
      mkAttr !value = LT.Attr ssNamespace { L.qName = name } (toString value)

instance ToElement T.Cell where
  toElement !c = emptyCell
    { L.elContent = map (LT.Elem . toElement) (maybeToList (T.cellData c))
    , L.elAttribs = catMaybes
      [ toA T.cellStyleID     "StyleID"     id
      , toA T.cellFormula     "Formula"     showFormula
      , toA T.cellIndex       "Index"       show
      , toA T.cellHRef        "HRef"        id
      , toA T.cellMergeAcross "MergeAcross" show
      , toA T.cellMergeDown   "MergeDown"   show
      ] }
    where
    showFormula (T.Formula !f) = f
    toA :: (T.Cell -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf !name toString = mkAttr <$> fieldOf c
      where
        mkAttr !value = LT.Attr ssNamespace { L.qName = name } (toString value)
  
instance ToElement I.ExcelValue where
   toElement !ev = mkData ev

instance ToElement I.Styles where
  toElement !ss = L.blank_element
    { L.elName = namespace { L.qName = "Styles"}
    , L.elContent = map (LT.Elem . toElement) $ I.styles ss }
      
instance ToElement I.Style where
  toElement !s = L.blank_element
    { L.elName = namespace { L.qName = "Style"}
    , L.elContent = map LT.Elem $ catMaybes $  
      [ toElement <$> I.styleAlignment s
      , toElement <$> I.styleFont s
      , toElement <$> I.styleInterior s
      ]
    , L.elAttribs = [ LT.Attr ssNamespace { L.qName = "ID" }  (I.styleID s)]
    }
    
instance ToElement I.Font where
  toElement (I.Font name family size iB iI c ) = L.blank_element
    { L.elName = namespace { L.qName = "Font" }
    , L.elAttribs = catMaybes
      [ LT.Attr ssNamespace { L.qName = "FontName" }             <$> name
      , LT.Attr xNamespace  { L.qName = "Family" }               <$> family
      , LT.Attr ssNamespace { L.qName = "Size" }   . show        <$> size
      , LT.Attr ssNamespace { L.qName = "Bold" }   . showBoolean <$> iB
      , LT.Attr ssNamespace { L.qName = "Italic" } . showBoolean <$> iI
      , LT.Attr ssNamespace { L.qName = "Color" }  . sRGB24show  <$> c
      ] }
        
instance ToElement I.Interior where
  toElement (I.Interior p c pc) = L.blank_element
    { L.elName = namespace { L.qName = "Interior" }
    , L.elAttribs = catMaybes
      [ LT.Attr ssNamespace { L.qName = "Color" } . sRGB24show        <$> c
      , LT.Attr ssNamespace { L.qName = "Pattern" }                   <$> p
      , LT.Attr ssNamespace { L.qName = "PatternColor" } . sRGB24show <$> pc
      ] }
    
instance ToElement I.Alignment where
  toElement align = L.blank_element
    { L.elName = namespace { L.qName = "Alignment" }
    , L.elAttribs = catMaybes
      [ LT.Attr ssNamespace { L.qName = "Horizontal" }             <$> I.alignmentHorizontal align
      , LT.Attr ssNamespace { L.qName = "Vertical" }               <$> I.alignmentVertical align
      , LT.Attr ssNamespace { L.qName = "WrapText" } . showBoolean <$> I.alignmentWrapText align
      ] }

