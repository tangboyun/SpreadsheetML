module Text.XML.SpreadsheetML.Builder where

import Text.XML.SpreadsheetML.Types
import qualified Text.XML.SpreadsheetML.Internal as I

-- | Construct empty values
emptyWorkbook :: Workbook
emptyWorkbook = Workbook Nothing (Just $ I.Styles []) []

emptyDocumentProperties :: DocumentProperties
emptyDocumentProperties =
  DocumentProperties Nothing Nothing Nothing Nothing Nothing Nothing Nothing

emptyWorksheet :: Name -> Worksheet
emptyWorksheet name = Worksheet Nothing name

emptyTable :: Table
emptyTable =
  Table [] [] Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing

emptyColumn :: Column
emptyColumn =
  Column Nothing Nothing Nothing Nothing Nothing Nothing Nothing

emptyRow :: Row
emptyRow = Row [] Nothing Nothing Nothing Nothing Nothing Nothing Nothing

emptyCell :: Cell
emptyCell = Cell Nothing Nothing Nothing Nothing Nothing Nothing

-- | Convenience constructors
number :: Double -> Cell
number d = emptyCell { cellData = Just (I.Number d) }

string :: String -> Cell
string s = emptyCell { cellData = Just (I.StringType s) }

bool :: Bool -> Cell
bool b = emptyCell { cellData = Just (I.Boolean b) }

richText :: String -> Cell
richText = undefined

-- | This function may change in future versions, if a real formula type is
-- created.
formula :: String -> Cell
formula f = emptyCell { cellFormula = Just (Formula f) }

mkWorkbook :: [Worksheet] -> Workbook
mkWorkbook ws = Workbook Nothing (Just $ I.Styles []) ws

mkWorksheet :: Name -> Table -> Worksheet
mkWorksheet name table = Worksheet (Just table) name

mkTable :: [Row] -> Table
mkTable rs = emptyTable { tableRows = rs }

mkRow :: [Cell] -> Row
mkRow cs = emptyRow { rowCells = cs }

-- | Most of the time this is the easiest way to make a table
tableFromCells :: [[Cell]] -> Table
tableFromCells cs = mkTable (map mkRow cs)

emptyStyle :: Style
emptyStyle = Style Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing

addStyle :: Name -> Style -> Workbook -> Workbook
addStyle (Name str) s w =
  let font = I.Font
             (fontName s)
             (fontFamily s)
             (fontSize s)
             (fontIsBold s)
             (fontIsItalic s)
             (fontColor s)
      s' = I.Style
           { I.styleID = str
           , I.styleAlignment = Just $ I.Alignment (hAlign s) (vAlign s)
           , I.styleFont = Just font
           , I.styleInterior = Just $ I.Interior Nothing (bgColor s) Nothing
           }
      ss = fmap (I.Styles . (++ [s']) . I.styles) (worksheetStyles w)
  in w { worksheetStyles = ss }

class StyleID a where
  withStyleID :: String -> a -> a
instance StyleID Cell where
  withStyleID str cell = cell { cellStyleID   = Just str }
instance StyleID Row where
  withStyleID str row  = row  { rowStyleID    = Just str  }
instance StyleID Column where
  withStyleID str col  = col  { columnStyleID = Just str }
instance StyleID Table where
  withStyleID str tab  = tab  { tableStyleID  = Just str }
  
class Format a where
  (#) :: b -> (b -> a) -> a
  (#) = flip ($)

instance Format Workbook
instance Format Table
instance Format Row
instance Format Column
instance Format Cell
instance Format I.ExcelValue

-- font :: Font -> a
-- align :: Align -> a
-- bgColor :: Colour b -> a
  
