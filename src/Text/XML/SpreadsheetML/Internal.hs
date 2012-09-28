-----------------------------------------------------------------------------
-- |
-- Module : 
-- Copyright : (c) 2012 Boyun Tang
-- License : BSD-style
-- Maintainer : tangboyun@hotmail.com
-- Stability : experimental
-- Portability : ghc
--
-- 
--
-----------------------------------------------------------------------------
module Text.XML.SpreadsheetML.Internal where

import Data.Colour.SRGB 

newtype Styles = Styles {
  styles :: [Style]
  }

data Style = Style
  -- Attr
  { styleID        :: String -- ^ Unique String ID
  -- Element    
  , styleAlignment :: Maybe Alignment
  , styleFont      :: Maybe Font
  , styleInterior  :: Maybe Interior
  }

data Font = Font
  { fontName     :: Maybe String
  , fontFamily   :: Maybe String
  , fontSize     :: Maybe Double
  , fontIsBold   :: Maybe Bool
  , fontIsItalic :: Maybe Bool
  , fontColor   :: Maybe (Colour Double)
  }

data Interior = Interior
  { interiorPattern       :: Maybe String
  , interiorColor        :: Maybe (Colour Double)
  , interiorPatternColor :: Maybe (Colour Double)
  }
           
data Alignment = Alignment
  { alignmentHorizontal :: Maybe String
  , alignmentVertical   :: Maybe String
  }

data ExcelValue = Number Double -- add RichText
                | Boolean Bool
                | StringType String
--                | ExcelValue RichText

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
