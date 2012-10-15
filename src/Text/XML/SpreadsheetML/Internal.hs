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
import Data.Maybe
import qualified Data.Map as M
-- import qualified Data.Vector.Unboxed as UV
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
  , alignmentWrapText :: Maybe Bool
  }

data ExcelValue = Number Double -- add RichText
                | Boolean Bool
                | StringType String
                | ExcelValue RichText
                 deriving (Show) 
data RichText = RichText
  { richTextContent :: String
  , richTextStyles :: [Op]
  }
  deriving (Show)
data Rich = B
          | F (Maybe String) (Maybe String) (Maybe Double) -- face,color,size
          | I
          | Sub
          | Sup
          | U
          deriving (Eq,Show)       

type BegIdx = Int
type EndIdx = Int

data RichStyle = R Rich
               | RS Rich RichStyle
               deriving (Eq,Show)
                        
data Op = Op (BegIdx,EndIdx) RichStyle
          deriving (Show)
                   
instance Eq Op where
  (==) (Op p1 _) (Op p2 _) = p1 == p2
  
instance Ord Op where
  compare (Op p1 _) (Op p2 _) = compare p1 p2

data Format = HasStyle String RichStyle
            | NoStyle String


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

slice :: Int -> Int -> String -> String
slice beg len str =  take len $ drop beg str

fromRich :: RichText -> String
fromRich (RichText str os) =
  case os of
    [] -> escStr str
    ops' ->
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
          ps = toP 0 $ concatMap (\(a,b) -> [a,b]) rs ++ [length str]
          doNothing (a,b) = NoStyle $ slice a (b-a) str
          doRich (a,b) = HasStyle (slice a (b-a) str)
                         (sMap M.! (a,b))
          segs = zipWith ($) (cycle [doNothing,doRich]) ps
      in if overlap
         then error "Range for styles can not overlap."
         else concatMap fromFormat segs

raw_data str = LT.CData { LT.cdVerbatim = LT.CDataRaw, LT.cdData = str, LT.cdLine = Nothing }


    
mkElement :: String -> RichStyle -> LT.Element
mkElement str' rs = go rs
  where go (R r) =
          let str = escStr str'
          in LT.blank_element { LT.elName = toName r
                              , LT.elContent = [LT.Text $ raw_data str]
                              , LT.elAttribs = toAttr r
                              }
        go (RS r rs) = 
            LT.blank_element { LT.elName = toName r
                             , LT.elContent = [LT.Elem $ go rs]
                             , LT.elAttribs = toAttr r
                             }
toName :: Rich -> LT.QName
toName r =
  case r of
    F _ _ _ -> LT.blank_name { LT.qName = "Font" }
    _       -> LT.blank_name { LT.qName = show r }

toAttr :: Rich -> [LT.Attr]
toAttr r =
  case r of
    F face color size ->
      catMaybes
      [fmap (mkAttr "Face") face
      ,fmap (mkAttr "Color") color
      ,fmap ((mkAttr "Size") . show) size
      ]
    _                 -> []  
  where mkAttr name prop =
          LT.Attr { LT.attrKey = LT.blank_name { LT.qName = name, LT.qPrefix = Just "html" }
                  , LT.attrVal = prop
                  }


fromFormat :: Format -> String
fromFormat (NoStyle str) = escStr str
fromFormat (HasStyle str rStyle) = showElement $
                                   mkElement str rStyle

