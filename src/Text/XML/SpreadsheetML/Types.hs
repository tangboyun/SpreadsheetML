module Text.XML.SpreadsheetML.Types where

{- See http://msdn.microsoft.com/en-us/library/aa140066%28office.10%29.aspx
 forked from https://github.com/dagit/SpreadsheetML
   Add more features.
-}
import Data.Word ( Word64 )
import Data.Colour
-- | Only implement what we need

data Workbook = Workbook
  { workbookDocumentProperties :: Maybe DocumentProperties
  , worksheetStyles      :: Maybe Styles -- ^ Add styles
  , workbookWorksheets         :: [Worksheet]
  }
  deriving (Read, Show)

newtype Styles = Styles {
  styles :: [Style]
  }
  deriving (Read, Show)               

data DocumentProperties = DocumentProperties
  { documentPropertiesTitle       :: Maybe String
  , documentPropertiesSubject     :: Maybe String
  , documentPropertiesKeywords    :: Maybe String
  , documentPropertiesDescription :: Maybe String
  , documentPropertiesRevision    :: Maybe Word64
  , documentPropertiesAppName     :: Maybe String
  , documentPropertiesCreated     :: Maybe String -- ^ Actually, this should be a date time
  }
  deriving (Read, Show)

data Worksheet = Worksheet
  { worksheetTable       :: Maybe Table
  , worksheetName        :: Name
  }
  deriving (Read, Show)

data Style = Style
  -- Attr
  { styleID        :: String -- ^ Unique String ID
  -- Element    
  , styleAlignment :: Maybe Alignment
  , styleFont      :: Maybe Font
  , styleInterior  :: Maybe Interior
  }
  deriving (Read, Show)

data Font = Font
  { fontName     :: Maybe String
  , fontFamily   :: Maybe String
  , fontColor    :: Maybe (Colour Double)
  , fontSize     :: Maybe Double
  , fontIsBold   :: Maybe Bool
  , fontIsItalic :: Maybe Bool
  }
  deriving (Read, Show)

data Interior = Interior
  { interiorColor        :: Colour Double
  , interiorPattern      :: String
  , interiorPatternColor :: Colour Double 
  }
  deriving (Read, Show)
           
data Alignment = Alignment
  { alignmentHorizontal :: Maybe Align
  , alignmentVertical   :: Maybe Align
  }
  deriving (Read, Show)


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
  }
  deriving (Read, Show)

data Column = Column
  { columnCaption      :: Maybe Caption
  , columnAutoFitWidth :: Maybe AutoFitWidth
  , columnHidden       :: Maybe Hidden
  , columnIndex        :: Maybe Word64
  , columnSpan         :: Maybe Word64
  , columnWidth        :: Maybe Double
  }
  deriving (Read, Show)

data Row = Row
  { rowCells         :: [Cell]
  , rowCaption       :: Maybe Caption
  , rowAutoFitHeight :: Maybe AutoFitHeight
  , rowHeight        :: Maybe Double
  , rowHidden        :: Maybe Hidden
  , rowIndex         :: Maybe Word64
  , rowSpan          :: Maybe Word64
  }
  deriving (Read, Show)

data Cell = Cell
  -- elements
  { cellData          :: Maybe ExcelValue
  -- Attributes
  , cellFormula       :: Maybe Formula
  , cellIndex         :: Maybe Word64
  , cellMergeAcross   :: Maybe Word64
  , cellMergeDown     :: Maybe Word64
  }
  deriving (Read, Show)

data ExcelValue = Number Double -- add RichText
                | Boolean Bool
                | StringType String
                | ExcelValue RichText
  deriving (Read, Show)

data RichText = RichText
  -- rich-text <ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">
  { richTextContent :: String
  -- Attributes  B, Font, I, S, Span, Sub, Sup, U
  , richTextFontTag :: Maybe [(Font,(Int,Int))]
  , richTextBTag    :: Maybe [(Int,Int)]
  , richTextITag    :: Maybe [(Int,Int)]
  , richTextSTag    :: Maybe [(Int,Int)]
  , richTextSpanTag :: Maybe [(Int,Int)]
  , richTextSubTag  :: Maybe [(Int,Int)]
  , richTextSupTag  :: Maybe [(Int,Int)]
  , richTextUTag    :: Maybe [(Int,Int)]
  }
  deriving (Read, Show)

-- | TODO: Currently just a string, but we could model excel formulas and
-- use that type here instead.
newtype Formula = Formula String
  deriving (Read, Show)

data AutoFitWidth = AutoFitWidth | DoNotAutoFitWidth
  deriving (Read, Show)

data AutoFitHeight = AutoFitHeight | DoNotAutoFitHeight
  deriving (Read, Show)

-- | Attribute for hidden things
data Hidden = Shown | Hidden
  deriving (Read, Show)

-- | For now this is just a string, but we could model excel's names
newtype Name = Name String
  deriving (Read, Show)

newtype Caption = Caption String
  deriving (Read, Show)

newtype Align = Align String
  deriving (Read, Show)
