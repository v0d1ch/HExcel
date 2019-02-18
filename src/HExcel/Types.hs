{-# LANGUAGE TemplateHaskell #-}

-- |
-- Module      :  HExcel.Types
-- Maintainer  :  Sasha Bogicevic <sasa.bogicevic@pm.me>
-- Stability   :  experimental
--
-- HExcel types are defined here
module HExcel.Types where

import Data.Time
import Data.Word
import Foreign
import Foreign.C.String
import Foreign.C.Types
import Lens.Micro.TH

-- | Excel Row
type Row = Word32
-- | Excel Column
type Col = Word16

-- | Pointer to Excel Workbook
data LxwWorkbook_
-- | Excel Workbook
newtype Workbook = Workbook (Ptr LxwWorkbook_)

data LxwWorksheet_
-- | Excel WorkSheet
newtype Worksheet = Worksheet (Ptr LxwWorksheet_)

-- | Pointer to Excel Format
data LxwFormat_
-- | Excel Format
newtype Format = Format { unFormat :: Ptr LxwFormat_ }

-- | 'HExcelState' is the state we thread trough for the writeCell function of the
-- HExcel typeclass. We are trying to create a convenient api for writing cell values
-- without too much hassle.
data HExcelState =
  HExcelState
    { _hexcelStateFormat     :: Maybe Format
      -- ^ define a formatting to be used on all cells
    , _hexcelStateSheet      :: Worksheet
      -- ^ define a worksheet to write to
    , _hexcelStateRowCeiling :: Word32
      -- ^ max row size
    , _hexcelStateInitRow    :: Row
      -- ^ first row to write to
    , _hexcelStateInitCol    :: Col
      -- ^ first column to write to
    , _hexcelStateRow        :: Row
      -- ^ current row
    , _hexcelStateCol        :: Col
      -- ^ current column
    }

makeLenses ''HExcelState

-- | Excel WorkBook Options
newtype WorkbookOptions =
  WorkbookOptions { workbookOptionsConstantMem :: Bool }

instance Storable WorkbookOptions where
  sizeOf _ = sizeOf (undefined :: Word8)
  alignment _ = alignment (undefined :: Int)
  peek ptr = do
    cm <- peekByteOff ptr 0 :: IO Word8
    return (WorkbookOptions (cm /= 1))
  poke ptr opts = do
    let cm | workbookOptionsConstantMem opts = 1
           | otherwise = 0
    pokeByteOff ptr 0 (cm :: Word8)

-- | Colors
data Color
  = ColorBlack
  | ColorBlue
  | ColorBrown
  | ColorCyan
  | ColorGray
  | ColorGreen
  | ColorLime
  | ColorMagenta
  | ColorNavy
  | ColorOrange
  | ColorPink
  | ColorPurple
  | ColorRed
  | ColorSilver
  | ColorWhite
  | ColorYellow
  | Color Word8 Word8 Word8

-- | Underline styles
data UnderlineStyle
  = UnderlineNone
  | UnderlineSingle
  | UnderlineDouble
  | UnderlineSingleAccounting
  | UnderlineDoubleAccounting
  deriving (Eq, Enum, Read, Show)

-- | Script styles
data ScriptStyle
  = SuperScript
  | SubScript
  deriving (Eq, Enum, Read, Show)

-- | Alignment styles
data Align
  = AlignNone
  | AlignLeft
  | AlignCenter
  | AlignRight
  | AlignFill
  | AlignJustify
  | AlignCenterAcross
  | AlignDistributed
  deriving (Eq, Enum, Read, Show)

-- | Vertical align styles
data VerticalAlign
  = VerticalAlignNone
  | VerticalAlignTop
  | VerticalAlignBottom
  | VerticalAlignCenter
  | VerticalAlignJustify
  | VerticalAlignDistributed
  deriving (Eq, Enum, Read, Show)

-- | Pattern styles
data Pattern
  = PatternNone
  | PatternSolid
  | PatternMediumGray
  | PatternDarkGray
  | PatternLightGray
  | PatternDarkHorizontal
  | PatternDarkVertical
  | PatternDarkDown
  | PatternDarkUp
  | PatternDarkGrid
  | PatternDarkTrellis
  | PatternLightHorizontal
  | PatternLightVertical
  | PatternLightDown
  | PatternLightUp
  | PatternLightGrid
  | PatternLightTrellis
  | PatternGray125
  | PatternGray0625
  deriving (Eq, Enum, Read, Show)

-- | Border options
data Border
  = BorderAll
  | BorderBottom
  | BorderTop
  | BorderLeft
  | BorderRight
  deriving (Eq, Read, Show)

-- | Border styles
data BorderStyle
  = BorderNone
  | BorderThin
  | BorderMedium
  | BorderDashed
  | BorderDotted
  | BorderThick
  | BorderDouble
  | BorderHair
  | BorderMediumDashed
  | BorderDashDot
  | BorderMediumDashDot
  | BorderDashDotDot
  | BorderMediumDashDotDot
  | BorderSlantDashDot
  deriving (Eq, Enum, Read, Show)

-- | Excel document properties
data DocProperties =
  DocProperties { docPropertiesTitle         :: String
                , docPropertiesSubject       :: String
                , docPropertiesAuthor        :: String
                , docPropertiesManager       :: String
                , docPropertiesCompany       :: String
                , docPropertiesCategory      :: String
                , docPropertiesKeywords      :: String
                , docPropertiesComments      :: String
                , docPropertiesStatus        :: String
                , docPropertiesHyperlinkBase :: String
                , docPropertiesCreated       :: UTCTime {-CTime-}
               }

data DocProperties' =
  DocProperties' { docPropsTitle         :: CString
                 , docPropsSubject       :: CString
                 , docPropsAuthor        :: CString
                 , docPropsManager       :: CString
                 , docPropsCompany       :: CString
                 , docPropsCategory      :: CString
                 , docPropsKeywords      :: CString
                 , docPropsComments      :: CString
                 , docPropsStatus        :: CString
                 , docPropsHyperlinkBase :: CString
                 , docPropsCreated       :: CTime
                 }

instance Storable DocProperties' where
  sizeOf _ = 10 * sizeOf (undefined :: CString) +
                  sizeOf (undefined :: CTime)
  alignment _ = alignment (undefined :: CString)
  peek = error "No implementation of 'peek' for 'DocProperties'"
  poke ptr props = do
    let n = sizeOf (undefined :: CString)
    pokeByteOff ptr (0 * n) (docPropsTitle props)
    pokeByteOff ptr (1 * n) (docPropsSubject props)
    pokeByteOff ptr (2 * n) (docPropsAuthor props)
    pokeByteOff ptr (3 * n) (docPropsManager props)
    pokeByteOff ptr (4 * n) (docPropsCompany props)
    pokeByteOff ptr (5 * n) (docPropsCategory props)
    pokeByteOff ptr (6 * n) (docPropsKeywords props)
    pokeByteOff ptr (7 * n) (docPropsComments props)
    pokeByteOff ptr (8 * n) (docPropsStatus props)
    pokeByteOff ptr (9 * n) (docPropsHyperlinkBase props)
    pokeByteOff ptr (10 * n) (docPropsCreated props)

-- | Type to hold datetime values
data DateTime =
  DateTime { dtYear   :: CInt
           , dtMonth  :: CInt
           , dtDay    :: CInt
           , dtHour   :: CInt
           , dtMinute :: CInt
           , dtSecond :: CDouble
           }
  deriving (Show)

instance Storable DateTime where
  sizeOf _ = 5 * sizeOf (undefined :: Int) +
                 sizeOf (undefined :: Double)
  alignment _ = alignment (undefined :: Int)
  peek ptr = do
    let ptr' = castPtr ptr
    DateTime <$> peekElemOff ptr' 0
             <*> peekElemOff ptr' 1
             <*> peekElemOff ptr' 2
             <*> peekElemOff ptr' 3
             <*> peekElemOff ptr' 4
             <*> peekElemOff ptr' 5
  poke ptr (DateTime y m d h mi s) = do
    pokeByteOff ptr 0 y
    pokeByteOff ptr 4 m
    pokeByteOff ptr 8 d
    pokeByteOff ptr 12 h
    pokeByteOff ptr 16 mi
    pokeByteOff ptr 20 s

-- | Type to hold image options
data ImageOptions =
  ImageOptions { imageOffsetX :: Int32
               , imageOffsetY :: Int32
               , imageScaleX  :: Double
               , imageScaleY  :: Double
               }

instance Storable ImageOptions where
  sizeOf _ = 2 * sizeOf (undefined :: Int32) +
             2 * sizeOf (undefined :: Double)
  alignment _ = alignment (undefined :: Int32)
  peek ptr =
    ImageOptions <$> peekByteOff ptr 0
                 <*> peekByteOff ptr 4
                 <*> peekByteOff ptr 8
                 <*> peekByteOff ptr 16
  poke ptr (ImageOptions ox oy sx sy) = do
    pokeByteOff ptr 0 ox
    pokeByteOff ptr 4 oy
    pokeByteOff ptr 8 sx
    pokeByteOff ptr 16 sy

-- | Paper size
data PaperSize
  = DefaultPaper
  | LetterPaper
  | A3Paper
  | A4Paper
  | A5Paper
  | OtherPaper Word8
  deriving (Eq)
