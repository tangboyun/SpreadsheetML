module Main where
import Text.XML.SpreadsheetML.Types
import Text.XML.SpreadsheetML.Writer
import Text.XML.SpreadsheetML.Builder
import Data.Colour.Names


rows1 = [ mkRow [ string "Group A" # mergeAcross 2 # withStyleID "g1"
                , string "Group B" # mergeAcross 2 # withStyleID "g2"
                , string "Fold Change" # withStyleID "bold"
                ]
        , mkRow [number 1,   number 1.2, number 0.9,
                 number 0.5 , number 0.3 ,number 0.7,
                 formula "=SUM(RC[-6]:RC[-4])/SUM(RC[-3]:RC[-1])"]
        , mkRow [number 1.4, number 1.1, number 0.8,
                 number 0.8, number 0.4, number 0.3,
                 formula "=SUM(RC[-6]:RC[-4])/SUM(RC[-3]:RC[-1])"]
        ]

worksheet1 = mkWorksheet (Name "Quantity Product Sheet") (mkTable rows1)
workbook = mkWorkbook [worksheet1]
           # addStyle (Name "Default") emptyStyle {fontName = Just "Times New Roman"} 
           # addStyle (Name "g1") emptyStyle { fontIsBold = Just True
                                             , bgColor = Just red
                                             , hAlign = Just "Center"
                                             }
           # addStyle (Name "g2") emptyStyle { fontIsBold = Just True
                                             , bgColor = Just green
                                             , hAlign = Just "Center"
                                             }
           # addStyle (Name "bold") emptyStyle { fontIsBold = Just True }

main :: IO ()
main = writeFile "test.xls" (showSpreadsheet workbook)
