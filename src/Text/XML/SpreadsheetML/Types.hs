module Text.XML.SpreadsheetML.Types where

{- See http://msdn.microsoft.com/en-us/library/aa140066%28office.10%29.aspx
   forked from https://github.com/dagit/SpreadsheetML
   Add more features.
-}
import Data.Word ( Word64 )
import Data.Colour

import qualified Text.XML.SpreadsheetML.Internal as I
-- | Only implement what we need

data Workbook = Workbook
  { workbookDocumentProperties :: Maybe DocumentProperties
  , worksheetStyles      :: Maybe I.Styles 
  , workbookWorksheets         :: [Worksheet]
  }

data Style = Style
  { fontName :: Maybe String
  , fontFamily :: Maybe String
  , fontColor :: Maybe (Colour Double)
  , fontSize :: Maybe Double
  , fontIsBold :: Maybe Bool
  , fontIsItalic :: Maybe Bool
  , hAlign :: Maybe String
  , vAlign :: Maybe String
  , wrapText :: Maybe Bool
  , bgColor :: Maybe (Colour Double)
  }

data DocumentProperties = DocumentProperties
  { documentPropertiesTitle       :: Maybe String
  , documentPropertiesSubject     :: Maybe String
  , documentPropertiesKeywords    :: Maybe String
  , documentPropertiesDescription :: Maybe String
  , documentPropertiesRevision    :: Maybe Word64
  , documentPropertiesAppName     :: Maybe String
  , documentPropertiesCreated     :: Maybe String
  }

data Worksheet = Worksheet
  { worksheetTable       :: Maybe Table
  , worksheetName        :: Name
  }



data Table = Table
  { tableColumns             :: [Column]
  , tableRows                :: [Row]
  , tableDefaultColumnWidth  :: Maybe Double -- ^ Default is 48
  , tableDefaultRowHeight    :: Maybe Double -- ^ Default is 12.75
  , tableExpandedColumnCount :: Maybe Word64
  , tableExpandedRowCount    :: Maybe Word64
  , tableLeftCell            :: Maybe Word64 -- ^ Default is 1
  , tableTopCell             :: Maybe Word64 -- ^ Default is 1
  , tableFullColumns         :: Maybe Bool
  , tableFullRows            :: Maybe Bool
  , tableStyleID             :: Maybe String
  }

data Column = Column
  { columnCaption      :: Maybe Caption
  , columnAutoFitWidth :: Maybe AutoFitWidth
  , columnHidden       :: Maybe Hidden
  , columnIndex        :: Maybe Word64
  , columnSpan         :: Maybe Word64
  , columnWidth        :: Maybe Double
  , columnStyleID      :: Maybe String
  }

data Row = Row
  { rowCells         :: [Cell]
  , rowCaption       :: Maybe Caption
  , rowAutoFitHeight :: Maybe AutoFitHeight
  , rowHeight        :: Maybe Double
  , rowHidden        :: Maybe Hidden
  , rowIndex         :: Maybe Word64
  , rowSpan          :: Maybe Word64
  , rowStyleID       :: Maybe String
  }

data Cell = Cell
  -- elements
  { cellData          :: Maybe I.ExcelValue
  -- Attributes
  , cellHRef          :: Maybe String
  , cellStyleID       :: Maybe String
  , cellFormula       :: Maybe Formula
  , cellIndex         :: Maybe Word64
  , cellMergeAcross   :: Maybe Word64
  , cellMergeDown     :: Maybe Word64
  }


newtype Formula = Formula String

data AutoFitWidth = AutoFitWidth | DoNotAutoFitWidth

data AutoFitHeight = AutoFitHeight | DoNotAutoFitHeight

data Hidden = Shown | Hidden

newtype Name = Name String

newtype Caption = Caption String


