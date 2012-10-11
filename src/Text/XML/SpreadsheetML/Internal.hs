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
import Data.List
import Data.Char
import qualified Data.Map as M
import qualified Data.Vector.Unboxed as UV
import qualified Text.XML.Light.Types as LT
import Text.XML.Light.Output (showElement)

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
                | ExcelValue RichText

data RichText = RichText
  -- rich-text <ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">
  { richTextContent :: String
  -- Attributes  B, Font, I, S, Span, Sub, Sup, U
  , richTextStyles :: Maybe [Op]
  }

data Rich = B
          | F String String Double -- face,color,size
          | I
          | S
          | Span
          | Sub
          | Sup
          | U
          deriving (Eq,Show)       

type BegIdx = Int
type EndIdx = Int

data RichStyle = R Rich
               | RS Rich RichStyle
               deriving (Eq)
                        
data Op = Op (BegIdx,EndIdx) RichStyle

instance Eq Op where
  (==) (Op p1 _) (Op p2 _) = p1 == p2
  
instance Ord Op where
  compare (Op p1 _) (Op p2 _) = compare p1 p2

data Format = HasStyle (UV.Vector Char) RichStyle
            | NoStyle (UV.Vector Char)


escStr cs = foldr escChar "" cs
escChar c = case c of
  '<'   -> showString "&lt;"
  '>'   -> showString "&gt;"
  '&'   -> showString "&amp;"
  '"'   -> showString "&quot;"
  '\''  -> showString "&#39;"
  _ | isPrint c -> showChar c
    | otherwise -> showString "&#" . shows oc . showChar ';'
    where oc = ord c

fromRich :: RichText -> String
fromRich (RichText str os) =
  case os of
    Nothing -> escStr str
    Just ops' ->
      let ops = nub $ sort ops'
          ipp = map (\(Op a b) -> (a,b)) ops
          rs = map fst ipp
          overlap = not $! fst $
                    foldl' (\(bool,(a,b)) (c,d) ->
                             let bool' = bool && (c >= b) && (b >= a)
                             in (bool',(c,d))
                           ) (True, head rs) (tail rs)
          toP _ [] = []
          toP x (y:ys) = (x,y):toP y ys
          sMap = M.fromList ipp
          ps = toP 0 $ concatMap (\(a,b) -> [a,b]) rs
          vec = UV.fromList str
          doNothing (a,b) = NoStyle $ UV.slice a (b-a) vec
          doRich (a,b) = HasStyle (UV.slice a (b-a) vec)
                         (sMap M.! (a,b))
          segs = zipWith ($) (cycle [doNothing,doRich]) ps
      in if overlap
         then error "Range for styles can not overlap."
         else concatMap fromFormat segs

raw_data str = LT.CData { LT.cdVerbatim = LT.CDataRaw, LT.cdData = str, LT.cdLine = Nothing }


    
mkElement :: UV.Vector Char -> RichStyle -> LT.Element
mkElement vec rs = go rs
  where go (R r) =
          let str = escStr $ UV.toList vec
          in LT.blank_element { LT.elName = LT.blank_name {LT.qName = toName r}
                              , LT.elContent = [LT.Text $ raw_data str]
                              , LT.elAttribs = toAttr r
                              }
        go (RS r rs) = 
            LT.blank_element { LT.elName = LT.blank_name {LT.qName = toName r}
                             , LT.elContent = [LT.Elem $ go rs]
                             , LT.elAttribs = toAttr r
                             }

toName = undefined
toAttr = undefined
                         


fromFormat :: Format -> String
fromFormat (NoStyle vec) = escStr $ UV.toList vec
fromFormat (HasStyle vec rStyle) = showElement $
                                   mkElement vec rStyle

