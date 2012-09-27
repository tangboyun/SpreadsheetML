{-# LANGUAGE ExistentialQuantification #-}
module Text.XML.SpreadsheetML.Types where

{- See http://msdn.microsoft.com/en-us/library/aa140066%28office.10%29.aspx
   forked from https://github.com/dagit/SpreadsheetML
   Add more features.
-}
import Data.Word ( Word64 )
import Data.Colour.SRGB

-- | Only implement what we need

data Workbook = Workbook
  { workbookDocumentProperties :: Maybe DocumentProperties
  , worksheetStyles      :: Maybe Styles -- ^ Add styles
  , workbookWorksheets         :: [Worksheet]
  }

newtype Styles = Styles {
  styles :: [Style]
  }

data DocumentProperties = DocumentProperties
  { documentPropertiesTitle       :: Maybe String
  , documentPropertiesSubject     :: Maybe String
  , documentPropertiesKeywords    :: Maybe String
  , documentPropertiesDescription :: Maybe String
  , documentPropertiesRevision    :: Maybe Word64
  , documentPropertiesAppName     :: Maybe String
  , documentPropertiesCreated     :: Maybe String -- ^ Actually, this should be a date time
  }

data Worksheet = Worksheet
  { worksheetTable       :: Maybe Table
  , worksheetName        :: Name
  }

data Style = Style
  -- Attr
  { styleID        :: String -- ^ Unique String ID
  -- Element    
  , styleAlignment :: Maybe Alignment
  , styleFont      :: Maybe Font
  , styleInterior  :: Maybe Interior
  }

data Font = forall a. Num a => Font
  { fontName     :: Maybe String
  , fontFamily   :: Maybe String
  , fontSize     :: Maybe Double
  , fontIsBold   :: Maybe Bool
  , fontIsItalic :: Maybe Bool
  , _fontColor   :: Maybe (Colour a)
  , _fshowColor  :: Colour a -> String
  }

data Interior = forall a. Num a => Interior
  { interiorPattern       :: Maybe String
  , _interiorColor        :: Maybe (Colour a)
  , _interiorPatternColor :: Maybe (Colour a)
  , _ishowColor           :: (Colour a) -> String
  }
           
data Alignment = Alignment
  { alignmentHorizontal :: Maybe String
  , alignmentVertical   :: Maybe String
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
  , columnStyleID         :: Maybe String
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
  { cellData          :: Maybe ExcelValue
  -- Attributes
  , cellStyleID       :: Maybe String
  , cellFormula       :: Maybe Formula
  , cellIndex         :: Maybe Word64
  , cellMergeAcross   :: Maybe Word64
  , cellMergeDown     :: Maybe Word64
  }

data ExcelValue = Number Double -- add RichText
                | Boolean Bool
                | StringType String
                | ExcelValue RichText

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

-- | TODO: Currently just a string, but we could model excel formulas and
-- use that type here instead.
newtype Formula = Formula String

data AutoFitWidth = AutoFitWidth | DoNotAutoFitWidth

data AutoFitHeight = AutoFitHeight | DoNotAutoFitHeight

-- | Attribute for hidden things
data Hidden = Shown | Hidden

-- | For now this is just a string, but we could model excel's names
newtype Name = Name String

newtype Caption = Caption String

