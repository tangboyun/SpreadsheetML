module Text.XML.SpreadsheetML.Builder where

import Text.XML.SpreadsheetML.Types
import Data.Word (Word64)
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
emptyCell = Cell Nothing Nothing Nothing Nothing Nothing Nothing Nothing

-- | Convenience constructors
number :: Double -> Cell
number d = emptyCell { cellData = Just (I.Number d) }

string :: String -> Cell
string s = emptyCell { cellData = Just (I.StringType s) }

bool :: Bool -> Cell
bool b = emptyCell { cellData = Just (I.Boolean b) }

  
-- | This function may change in future versions, if a real formula type is
-- created.
formula :: String -> Cell
formula f = emptyCell { cellFormula = Just (Formula f) }

href :: String -> String -> Cell
href url showStr = emptyCell { cellData = Just (I.StringType showStr)
                             , cellHRef = Just url
                             , cellFormula = Just (Formula $ "=HYPERLINK("++ show url++","++ show showStr++")")
                             }

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
emptyStyle = Style Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing


mergeDown :: Word64 -> Cell -> Cell
mergeDown n c = c { cellMergeDown = Just n }
  
mergeAcross :: Word64 -> Cell -> Cell
mergeAcross n c = c { cellMergeAcross = Just n }


  
