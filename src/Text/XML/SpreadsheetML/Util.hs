{-# LANGUAGE BangPatterns #-}
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

module Text.XML.SpreadsheetML.Util
       (
         TextProperty(..)
       , FontProperty(..)
       , dfp
       , addStyle
       , addTextPropertyAtRanges
       , Index(..)
       , StyleID(..)
       , Format(..)
       )
       where

import Data.Word (Word64)
import Text.XML.SpreadsheetML.Types 
import qualified Text.XML.SpreadsheetML.Internal as I
import           Data.Colour.SRGB
import Control.Applicative
import Data.List
data TextProperty = Bold
                  | Italic
                  | UnderLine
                  | Subscript
                  | Superscript
                  | Text FontProperty
                    
data FontProperty = FP
  { name :: !(Maybe String) -- ^ Font name
  , color :: !(Maybe (Colour Double)) -- ^ Font color
  , size :: !(Maybe Double)           -- ^ Font size
  }


-- | default font property
dfp :: FontProperty
dfp = FP Nothing Nothing Nothing

addStyle :: Name -> Style -> Workbook -> Workbook
addStyle (Name !str) !s !w =
  let !font = I.Font
             (fontName s)
             (fontFamily s)
             (fontSize s)
             (fontIsBold s)
             (fontIsItalic s)
             (fontColor s)
      !s' = I.Style
           { I.styleID = str
           , I.styleAlignment = Just $ I.Alignment (hAlign s) (vAlign s) (wrapText s)
           , I.styleFont = Just font
           , I.styleInterior = case bgColor s of
                                    Nothing -> Nothing
                                    _       -> Just $ I.Interior (Just "Solid") (bgColor s) Nothing
           }
      !ss = fmap (I.Styles . (++ [s']) . I.styles) (worksheetStyles w)
  in w { worksheetStyles = ss }

addTextPropertyAtRanges :: [(Int,Int)] -> [TextProperty] -> Cell -> Cell
addTextPropertyAtRanges rs ts c = foldl' (\cell r -> add cell ts r) c rs

addTextPropertyAtRange :: (Int,Int) -> [TextProperty] -> Cell -> Cell
addTextPropertyAtRange range ts c = add c ts range

add :: Cell -> [TextProperty]-> (Int,Int) -> Cell
add !c !ts !range =
  let !r = buildRichStyle ts
      (!str,ops) = extract c
  in
   case r of
     Nothing -> c
     Just s ->
       c { cellData = Just $! I.ExcelValue $! I.RichText str $!
                      ops ++ [I.Op range s]
         }
  where 
    extract :: Cell -> (String, [I.Op])
    extract !c =
      case cellData c of
        Nothing -> ("",[])
        Just (I.StringType s) -> (s,[])
        Just (I.ExcelValue t) -> (I.richTextContent t, I.richTextStyles t)
        _            -> error "Not a STRING"
    buildRichStyle :: [TextProperty] -> Maybe I.RichStyle
    buildRichStyle [] = Nothing
    buildRichStyle (r:rs) = Just $ go (I.R $ propToRich r) rs
    go r [] = r
    go r1 (r2:rs) = go (I.RS  (propToRich r2) r1) rs
    propToRich :: TextProperty -> I.Rich
    propToRich !t =
      case t of
        Bold -> I.B
        Italic -> I.I
        UnderLine -> I.U
        Subscript -> I.Sub
        Superscript -> I.Sup
        Text (FP n c s) -> I.F n (fmap sRGB24show c) s
    

      
class Index a where
  begAtIdx :: Word64 -> a -> a
instance Index Cell where
  begAtIdx !n !cel = cel { cellIndex = Just n }
instance Index Row where
  begAtIdx !n !row = row { rowIndex = Just n }
instance Index Column where
  begAtIdx !n !col = col { columnIndex = Just n }
  

class StyleID a where
  withStyleID :: String -> a -> a
instance StyleID Cell where
  withStyleID !str !cell = cell { cellStyleID   = Just str }
instance StyleID Row where
  withStyleID !str !row  = row  { rowStyleID    = Just str  }
instance StyleID Column where
  withStyleID !str !col  = col  { columnStyleID = Just str }
instance StyleID Table where
  withStyleID !str !tab  = tab  { tableStyleID  = Just str }
  
class Format a where
  -- | It’s just reverse function application
  (#) :: b -> (b -> a) -> a 
  (#) = flip ($)

instance Format Workbook
instance Format Table
instance Format Row
instance Format Column
instance Format Cell
