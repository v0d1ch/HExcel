{-# LANGUAGE ConstrainedClassMethods #-}
{-# LANGUAGE FlexibleContexts        #-}
{-# LANGUAGE FlexibleInstances       #-}
{-# LANGUAGE InstanceSigs            #-}
{-# LANGUAGE RecordWildCards         #-}
{-# LANGUAGE ScopedTypeVariables     #-}

-- |
-- Module      :  HExcel.HExcelInternal
-- Maintainer  :  Sasha Bogicevic <sasa.bogicevic@pm.me>
-- Stability   :  experimental
--
-- This module contains almost all of the library functionality.

module HExcel.HExcelInternal
  ( Workbook
  , workbookNew
  , workbookNewConstantMem
  , workbookClose
  , workbookAddWorksheet
  , workbookAddFormat
  , workbookDefineName
  , DocProperties (..)
  , workbookSetProperties
  , Worksheet
  , Row
  , Col
  , writeNumber
  , writeString
  , writeUTCTime
  , writeFormula
  , writeArrayFormula
  , DateTime (..)
  , utcTimeToDateTime
  , zonedTimeToDateTime
  , writeDateTime
  , writeUrl
  , worksheetSetRow
  , worksheetSetColumn
  , ImageOptions (..)
  , worksheetInsertImage
  , worksheetInsertImageOpt
  , worksheetMergeRange
  , worksheetFreezePanes
  , worksheetSplitPanes
  , worksheetSetLandscape
  , worksheetSetPortrait
  , worksheetSetPageView
  , PaperSize (..)
  , worksheetSetPaperSize
  , worksheetSetMargins
  , worksheetSetHeaderCtl
  , worksheetSetFooterCtl
  , worksheetSetZoom
  , worksheetSetPrintScale
  , Format
  , formatSetFontName
  , formatSetFontSize
  , Color (..)
  , formatSetFontColor
  , formatSetNumFormat
  , formatSetBold
  , formatSetItalic
  , UnderlineStyle (..)
  , formatSetUnderline
  , formatSetStrikeout
  , ScriptStyle (..)
  , formatSetScript
  , formatSetBuiltInFormat
  , Align (..)
  , VerticalAlign (..)
  , formatSetAlign
  , formatSetVerticalAlign
  , formatSetTextWrap
  , formatSetRotation
  , formatSetShrink
  , Pattern (..)
  , formatSetPattern
  , formatSetBackgroundColor
  , formatSetForegroundColor
  , Border (..)
  , BorderStyle (..)
  , formatSetBorder
  , formatSetBorderColor
  , HExcelEnv (..)
  , HExcel (..)
  ) where

import Control.Monad.IO.Class
import Control.Monad.Trans.Reader
import Control.Monad.Trans.State
import Control.Monad.Trans.Class (lift)
import Data.Time
import Data.Time.Clock.POSIX
import Foreign
import Foreign.C.String
import Foreign.C.Types
import HExcel.Ffi
import HExcel.Types
import Data.Bifunctor

-- | HExcel class that provides a single function `writeCell` as a convenient method
-- of writing excel cell values
--
-- >  wb <- workbookNew "test.xlsx"
-- >  let props =
-- >        def { docPropertiesTitle   = "Test Workbook"
-- >            , docPropertiesCompany = "HExcel"
-- >            }
-- >  workbookSetProperties wb props
-- >  ws <- workbookAddWorksheet wb "First Sheet"
-- >  df <- workbookAddFormat wb
-- >  formatSetNumFormat df "mmm d yyyy hh:mm AM/PM"
-- >  now <- getZonedTime
-- >  -- You can create HExcelEnv if you plan to write rows and cols in a loop
-- >  let env = HExcelEnv (0 :: Row) (0 :: Col) 10 Nothing ws
-- >  flip execStateT (5 :: Row,5 :: Col) $ do
-- >    flip runReaderT env $ liftIO $ do
-- >      forM_ [0 .. 10] $ \n -> do
-- >        (row, col) <- lift get
-- >        -- you can use HExcel class funtion `writeCell` to write values to file
-- >        writeCell "David"
-- >        writeCell (3 :: Double)
-- >        writeCell "Dimi"
-- >        writeCell (0 :: Double)
-- >        -- or use specialized functions
-- >        writeString ws (Just df) 1 3 "xxx"
-- >        writeNumber ws (Just df) 2 4 "yyy"
-- >        writeDateTime ws (Just df) 3 5 (zonedTimeToDateTime now)
-- >  liftIO $ workbookClose wb
--
class HExcel a where
  writeCell :: MonadIO m => a -> ReaderT HExcelEnv (StateT (Row, Col) m) ()

instance HExcel String where
  writeCell :: MonadIO m => String -> ReaderT HExcelEnv (StateT (Row, Col) m) ()
  writeCell val = ask >>= \HExcelEnv {..} -> do
    rowAndCol@(row, col) <- lift get
    liftIO $ writeString hexcelEnvSheet hexcelEnvFormat row col val
    lift $ put (bimap (+1) (+1) rowAndCol)

instance HExcel Double where
  writeCell :: MonadIO m => Double -> ReaderT HExcelEnv (StateT (Row, Col) m) ()
  writeCell val = ask >>= \HExcelEnv {..} -> do
    rowAndCol@(row, col) <- lift get
    liftIO $ writeNumber hexcelEnvSheet hexcelEnvFormat row col val
    lift $ put (bimap (+1) (+1) rowAndCol)

instance HExcel UTCTime where
  writeCell :: MonadIO m => UTCTime -> ReaderT HExcelEnv (StateT (Row, Col) m) ()
  writeCell val = ask >>= \HExcelEnv {..} -> do
    rowAndCol@(row, col) <- lift get
    liftIO $ writeUTCTime hexcelEnvSheet hexcelEnvFormat row col val
    lift $ put (bimap (+1) (+1) rowAndCol)

instance HExcel DateTime where
  writeCell :: MonadIO m => DateTime -> ReaderT HExcelEnv (StateT (Row, Col) m) ()
  writeCell val = ask >>= \HExcelEnv {..} -> do
    rowAndCol@(row, col) <- lift get
    liftIO $ writeDateTime hexcelEnvSheet hexcelEnvFormat row col val
    lift $ put (bimap (+1) (+1) rowAndCol)

-- | Create new workbook
workbookNew :: FilePath -> IO Workbook
workbookNew path = withCString path $ fmap Workbook . workbook_new

-- | Create new workbook but force constant memory.
-- It reduces the amount of data stored in memory so that large files can be written efficiently.
workbookNewConstantMem :: FilePath -> IO Workbook
workbookNewConstantMem path =
  with (WorkbookOptions True) $ \copts ->
    withCString path $ \cpath ->
      Workbook <$> workbook_new_opt cpath copts

-- | Close the workbook
workbookClose :: Workbook -> IO ()
workbookClose (Workbook wb) =
  workbook_close wb

-- | Add the worksheet
workbookAddWorksheet :: Workbook -> String -> IO Worksheet
workbookAddWorksheet (Workbook wb) name =
  withCString name $ fmap Worksheet . workbook_add_worksheet wb

-- | Add workbook format
workbookAddFormat :: Workbook -> IO Format
workbookAddFormat (Workbook wb) =
  Format <$> workbook_add_format wb

withDocProperties :: DocProperties -> (Ptr DocProperties' -> IO a) -> IO a
withDocProperties props action =
  withCString (docPropertiesTitle props) $ \ctitle ->
  withCString (docPropertiesSubject props) $ \csubject ->
  withCString (docPropertiesAuthor props) $ \cauthor ->
  withCString (docPropertiesManager props) $ \cmanager ->
  withCString (docPropertiesCompany props) $ \ccompany ->
  withCString (docPropertiesCategory props) $ \ccat ->
  withCString (docPropertiesKeywords props) $ \ckws ->
  withCString (docPropertiesComments props) $ \ccmts ->
  withCString (docPropertiesStatus props) $ \cstat ->
  withCString (docPropertiesHyperlinkBase props) $ \clb ->
    let time   = CTime (round (utcTimeToPOSIXSeconds (docPropertiesCreated props)))
        props' = DocProperties' ctitle csubject cauthor cmanager
                                ccompany ccat ckws ccmts cstat clb time
    in with props' action

-- | Set workbook properties
workbookSetProperties :: Workbook -> DocProperties -> IO ()
workbookSetProperties (Workbook wb) props =
  withDocProperties props $ \cprops ->
    workbook_set_properties wb cprops

-- | Set workbook name
workbookDefineName :: Workbook -> String -> String -> IO ()
workbookDefineName (Workbook wb) name formula =
  withCString name $ \cname -> withCString formula $ \cformula ->
    workbook_define_name wb cname cformula

-- | Write a 'Double' value to Excel cell
writeNumber :: Worksheet -> Maybe Format -> Row -> Col -> Double -> IO ()
writeNumber (Worksheet ws) mfmt row col number =
  worksheet_write_number ws row col number (maybe nullPtr unFormat mfmt)

-- | Write a 'String' value to Excel cell
writeString :: Worksheet -> Maybe Format -> Row -> Col -> String -> IO ()
writeString (Worksheet ws) mfmt row col str =
  withCString str $ \cstr ->
    worksheet_write_string ws row col cstr (maybe nullPtr unFormat mfmt)

-- | Write a 'UTCTime' value to Excel cell
writeUTCTime :: Worksheet -> Maybe Format -> Row -> Col -> UTCTime -> IO ()
writeUTCTime (Worksheet ws) mfmt row col t = do
    let tz = localTimeToUTC utc . utcToLocalTime (TimeZone 60 True "BST")
        ts = utcTimeToPOSIXSeconds (read "1900-01-01 00:00:00") - (2 * 24 * 60 * 60)
        ft = fromRational . toRational . (/ (24 * 60 * 60)) . (+ negate ts) . utcTimeToPOSIXSeconds . tz
    worksheet_write_number ws row col (ft t) (maybe nullPtr unFormat mfmt)

-- | Write a formula to Excel cell
writeFormula :: Worksheet -> Maybe Format -> Row -> Col -> String -> IO ()
writeFormula (Worksheet ws) mfmt row col str =
  withCString str $ \cstr ->
    worksheet_write_formula ws row col cstr (maybe nullPtr unFormat mfmt)

writeArrayFormula
  :: Worksheet
  -> Maybe Format
  -> Row
  -> Col
  -> Row
  -> Col
  -> String
  -> IO ()
writeArrayFormula (Worksheet ws) mfmt frow fcol erow ecol str =
  withCString str $ \cstr ->
    worksheet_write_array_formula
      ws
      frow
      fcol
      erow
      ecol
      cstr
      (maybe nullPtr unFormat mfmt)

-- | Helper function to convert  'UTCTime' to 'DateTime'
utcTimeToDateTime :: UTCTime -> DateTime
utcTimeToDateTime (UTCTime day time) =
  let (y, m, d)        = toGregorian day
      TimeOfDay h mi s = timeToTimeOfDay time
  in DateTime (fromIntegral y) (fromIntegral m) (fromIntegral d)
       (fromIntegral h) (fromIntegral mi) (fromRational (toRational s))

-- | Helper function to convert  'ZonedTime' to 'DateTime'
zonedTimeToDateTime :: ZonedTime -> DateTime
zonedTimeToDateTime = utcTimeToDateTime . zonedTimeToUTC

-- | Write a 'DateTime' to Excel cell
writeDateTime :: Worksheet -> Maybe Format -> Row -> Col -> DateTime -> IO ()
writeDateTime (Worksheet ws) mfmt row col dt =
  with dt $ \pdt ->
    worksheet_write_datetime ws row col pdt (maybe nullPtr unFormat mfmt)

-- | Write a url to Excel cell
writeUrl :: Worksheet -> Maybe Format -> Row -> Col -> String -> IO ()
writeUrl (Worksheet ws) mfmt row col str =
  withCString str $ \cstr ->
    worksheet_write_url ws row col cstr (maybe nullPtr unFormat mfmt)

-- | Set worksheet row
worksheetSetRow :: Worksheet -> Maybe Format -> Row -> Double -> IO ()
worksheetSetRow (Worksheet ws) mfmt row height =
  worksheet_set_row ws row height (maybe nullPtr unFormat mfmt)

-- | Set worksheet column
worksheetSetColumn :: Worksheet -> Maybe Format -> Col -> Col -> Double -> IO ()
worksheetSetColumn (Worksheet ws) mfmt fcol lcol width =
  worksheet_set_column ws fcol lcol width (maybe nullPtr unFormat mfmt)

-- | Insert image to worksheet
worksheetInsertImage :: Worksheet -> Word32 -> Word16 -> String -> IO ()
worksheetInsertImage (Worksheet ws) row col path =
  withCString path $ \cpath ->
    worksheet_insert_image ws row col cpath

worksheetInsertImageOpt
  :: Worksheet
  -> Row
  -> Col
  -> FilePath
  -> ImageOptions
  -> IO ()
worksheetInsertImageOpt (Worksheet ws) row col path opt =
  withCString path $ \cpath ->
    with opt $ \optr -> worksheet_insert_image_opt ws row col cpath optr

-- | Merge columns
worksheetMergeRange
  :: Worksheet
  -> Maybe Format
  -> Row
  -> Col
  -> Row
  -> Col
  -> String
  -> IO ()
worksheetMergeRange (Worksheet ws) mfmt frow fcol lrow lcol str =
  withCString str $ \cstr ->
    worksheet_merge_range
      ws
      frow
      fcol
      lrow
      lcol
      cstr
      (maybe nullPtr unFormat mfmt)

worksheetFreezePanes :: Worksheet -> Row -> Col -> IO ()
worksheetFreezePanes (Worksheet ws) = worksheet_freeze_panes ws

worksheetSplitPanes :: Worksheet -> Double -> Double -> IO ()
worksheetSplitPanes (Worksheet ws) = worksheet_split_panes ws

-- | Set worksheet to Landscape
worksheetSetLandscape :: Worksheet -> IO ()
worksheetSetLandscape (Worksheet ws) =
  worksheet_set_landscape ws

-- | Set worksheet to Portrait
worksheetSetPortrait :: Worksheet -> IO ()
worksheetSetPortrait (Worksheet ws) =
  worksheet_set_portrait ws

worksheetSetPageView :: Worksheet -> IO ()
worksheetSetPageView (Worksheet ws) =
  worksheet_set_page_view ws

-- | Set worksheet 'PaperSize'
worksheetSetPaperSize :: Worksheet -> PaperSize -> IO ()
worksheetSetPaperSize (Worksheet ws) paper =
  worksheet_set_paper ws (toPaper paper)
  where
    toPaper :: PaperSize -> Word8
    toPaper DefaultPaper   = 0
    toPaper LetterPaper    = 1
    toPaper A3Paper        = 8
    toPaper A4Paper        = 9
    toPaper A5Paper        = 11
    toPaper (OtherPaper n) = n

-- | Set worksheet margins
worksheetSetMargins :: Worksheet -> Double -> Double -> Double -> Double -> IO ()
worksheetSetMargins (Worksheet ws) = worksheet_set_margins ws

worksheetSetHeaderCtl :: Worksheet -> String -> IO ()
worksheetSetHeaderCtl (Worksheet ws) str =
  withCString str $ \cstr -> worksheet_set_header ws cstr

worksheetSetFooterCtl :: Worksheet -> String -> IO ()
worksheetSetFooterCtl (Worksheet ws) str =
  withCString str $ \cstr -> worksheet_set_footer ws cstr

worksheetSetZoom :: Worksheet -> Double -> IO ()
worksheetSetZoom (Worksheet ws) zoom =
  worksheet_set_zoom ws (round (100.0 * zoom'))
  where
    zoom' = min 0.1 (max 4.0 zoom)

worksheetSetPrintScale :: Worksheet -> Double -> IO ()
worksheetSetPrintScale (Worksheet ws) scale =
  worksheet_set_print_scale ws (round (100.0 * scale'))
  where
    scale' = min 0.1 (max 4.0 scale)

-- | Set font name
formatSetFontName :: Format -> String -> IO ()
formatSetFontName (Format fp) name =
  withCString name $ \cname ->
    format_set_font_name fp cname

-- | Set font size
formatSetFontSize :: Format -> Word16 -> IO ()
formatSetFontSize (Format fp) = format_set_font_size fp

colorIndex :: Color -> Int32
colorIndex ColorBlack    = 0x00000000
colorIndex ColorBlue     = 0x000000ff
colorIndex ColorBrown    = 0x00800000
colorIndex ColorCyan     = 0x0000ffff
colorIndex ColorGray     = 0x00808080
colorIndex ColorGreen    = 0x00008000
colorIndex ColorLime     = 0x0000ff00
colorIndex ColorMagenta  = 0x00ff00ff
colorIndex ColorNavy     = 0x00000080
colorIndex ColorOrange   = 0x00ff6600
colorIndex ColorPink     = 0x00ff00ff
colorIndex ColorPurple   = 0x00800080
colorIndex ColorRed      = 0x00ff0000
colorIndex ColorSilver   = 0x00c0c0c0
colorIndex ColorWhite    = 0x00ffffff
colorIndex ColorYellow   = 0x00ffff00
colorIndex (Color r g b) =
  fromIntegral r `shiftL` 16 .|.
  fromIntegral g `shiftL`  8 .|.
  fromIntegral b

-- | Set font color
formatSetFontColor :: Format -> Color -> IO ()
formatSetFontColor (Format fp) color =
  format_set_font_color fp (colorIndex color)

-- | Set number format
formatSetNumFormat :: Format -> String -> IO ()
formatSetNumFormat (Format fp) fmt =
  withCString fmt $ \cfmt ->
    format_set_num_format fp cfmt

-- | Set bold style
formatSetBold :: Format -> IO ()
formatSetBold (Format fp) =
  format_set_bold fp

-- | Set italic style
formatSetItalic :: Format -> IO ()
formatSetItalic (Format fp) =
  format_set_italic fp

-- | Set underline style
formatSetUnderline :: Format -> UnderlineStyle -> IO ()
formatSetUnderline (Format fp) us =
  format_set_underline fp (fromIntegral (fromEnum us))

formatSetStrikeout :: Format -> IO ()
formatSetStrikeout (Format fp) =
  format_set_font_strikeout fp

formatSetScript :: Format -> ScriptStyle -> IO ()
formatSetScript (Format fp) s =
  format_set_font_script fp (1 + fromIntegral (fromEnum s))

formatSetBuiltInFormat :: Format -> Word8 -> IO ()
formatSetBuiltInFormat (Format fp) = format_set_num_format_index fp

formatSetAlign :: Format -> Align -> IO ()
formatSetAlign (Format fp) a = format_set_align fp (fromIntegral (fromEnum a))

formatSetVerticalAlign :: Format -> VerticalAlign -> IO ()
formatSetVerticalAlign (Format fp) a =
  format_set_align fp a'
  where
    a' = case fromEnum a of
      0 -> 0
      n -> 7 + fromIntegral n

formatSetTextWrap :: Format -> IO ()
formatSetTextWrap (Format fp) =
  format_set_text_wrap fp

formatSetRotation :: Format -> Int -> IO ()
formatSetRotation (Format fp) angle =
  format_set_rotation fp (fromIntegral angle)

formatSetShrink :: Format -> IO ()
formatSetShrink (Format fp) =
  format_set_shrink fp


formatSetPattern :: Format -> Pattern -> IO ()
formatSetPattern (Format fp) pat =
  format_set_pattern fp (fromIntegral (fromEnum pat))

formatSetBackgroundColor :: Format -> Color -> IO ()
formatSetBackgroundColor (Format fp) color =
  format_set_bg_color fp (colorIndex color)

formatSetForegroundColor :: Format -> Color -> IO ()
formatSetForegroundColor (Format fp) color =
  format_set_fg_color fp (colorIndex color)


formatSetBorder :: Format -> Border -> BorderStyle -> IO ()
formatSetBorder (Format fp) border style =
  function fp (fromIntegral (fromEnum style))
  where
    function = case border of
      BorderAll    -> format_set_border
      BorderBottom -> format_set_bottom
      BorderTop    -> format_set_top
      BorderLeft   -> format_set_left
      BorderRight  -> format_set_right

formatSetBorderColor :: Format -> Border -> Color -> IO ()
formatSetBorderColor (Format fp) border color =
  function fp (colorIndex color)
  where
    function = case border of
      BorderAll    -> format_set_border_color
      BorderBottom -> format_set_bottom_color
      BorderTop    -> format_set_top_color
      BorderLeft   -> format_set_left_color
      BorderRight  -> format_set_right_color
