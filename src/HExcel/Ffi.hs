{-# LANGUAGE ForeignFunctionInterface #-}

-- |
-- Module      :  HExcel.Ffi
-- Maintainer  :  Sasha Bogicevic <sasa.bogicevic@pm.me>
-- Stability   :  experimental
--
-- FFI code

module HExcel.Ffi where

import Foreign
import Foreign.C.String
import HExcel.Types

foreign import ccall "workbook_new"
  workbook_new :: CString -> IO (Ptr LxwWorkbook_)

foreign import ccall "workbook_new_opt"
  workbook_new_opt :: CString -> Ptr WorkbookOptions -> IO (Ptr LxwWorkbook_)

foreign import ccall "workbook_close"
  workbook_close :: Ptr LxwWorkbook_ -> IO ()

foreign import ccall "workbook_add_worksheet"
  workbook_add_worksheet :: Ptr LxwWorkbook_ -> CString ->
                            IO (Ptr LxwWorksheet_)

foreign import ccall "workbook_add_format"
  workbook_add_format :: Ptr LxwWorkbook_ -> IO (Ptr LxwFormat_)

foreign import ccall "workbook_set_properties"
  workbook_set_properties :: Ptr LxwWorkbook_ ->
                             Ptr DocProperties' -> IO ()

foreign import ccall "workbook_define_name"
  workbook_define_name :: Ptr LxwWorkbook_ -> CString -> CString -> IO ()

foreign import ccall "worksheet_write_number"
  worksheet_write_number :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 ->
                            Double -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_write_string"
  worksheet_write_string :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 ->
                            CString -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_write_formula"
  worksheet_write_formula :: Ptr LxwWorksheet_ ->
                             Word32 -> Word16 ->
                             CString -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_write_array_formula"
  worksheet_write_array_formula :: Ptr LxwWorksheet_ ->
                                   Word32 -> Word16 ->
                                   Word32 -> Word16 ->
                                   CString -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_write_datetime"
  worksheet_write_datetime :: Ptr LxwWorksheet_ ->
                              Word32 -> Word16 ->
                              Ptr DateTime -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_write_url"
  worksheet_write_url :: Ptr LxwWorksheet_ ->
                         Word32 -> Word16 ->
                         CString -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_set_row"
  worksheet_set_row :: Ptr LxwWorksheet_ ->
                       Word32 -> Double -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_set_column"
  worksheet_set_column :: Ptr LxwWorksheet_ ->
                          Word16 -> Word16 -> Double -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_insert_image"
  worksheet_insert_image :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 -> CString -> IO ()

foreign import ccall "worksheet_insert_image_opt"
  worksheet_insert_image_opt :: Ptr LxwWorksheet_ ->
                                Word32 -> Word16 ->
                                CString -> Ptr ImageOptions -> IO ()

foreign import ccall "worksheet_merge_range"
  worksheet_merge_range :: Ptr LxwWorksheet_ ->
                           Word32 -> Word16 ->
                           Word32 -> Word16 ->
                           CString -> Ptr LxwFormat_ -> IO ()

foreign import ccall "worksheet_freeze_panes"
  worksheet_freeze_panes :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 -> IO ()

foreign import ccall "worksheet_split_panes"
  worksheet_split_panes :: Ptr LxwWorksheet_ ->
                           Double -> Double -> IO ()

foreign import ccall "worksheet_set_landscape"
  worksheet_set_landscape :: Ptr LxwWorksheet_ -> IO ()

foreign import ccall "worksheet_set_portrait"
  worksheet_set_portrait :: Ptr LxwWorksheet_ -> IO ()

foreign import ccall "worksheet_set_page_view"
  worksheet_set_page_view :: Ptr LxwWorksheet_ -> IO ()

foreign import ccall "worksheet_set_paper"
  worksheet_set_paper :: Ptr LxwWorksheet_ -> Word8 -> IO ()

foreign import ccall "worksheet_set_margins"
  worksheet_set_margins :: Ptr LxwWorksheet_ ->
                           Double -> Double ->
                           Double -> Double -> IO ()

foreign import ccall "worksheet_set_header"
  worksheet_set_header :: Ptr LxwWorksheet_ -> CString -> IO ()

foreign import ccall "worksheet_set_footer"
  worksheet_set_footer :: Ptr LxwWorksheet_ -> CString -> IO ()

foreign import ccall "worksheet_set_zoom"
  worksheet_set_zoom :: Ptr LxwWorksheet_ -> Word16 -> IO ()

foreign import ccall "worksheet_set_print_scale"
  worksheet_set_print_scale:: Ptr LxwWorksheet_ -> Word16 -> IO ()

foreign import ccall "format_set_font_name"
  format_set_font_name :: Ptr LxwFormat_ -> CString -> IO ()

foreign import ccall "format_set_font_size"
  format_set_font_size :: Ptr LxwFormat_ -> Word16 -> IO ()

foreign import ccall "format_set_font_color"
  format_set_font_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_num_format"
  format_set_num_format :: Ptr LxwFormat_ -> CString -> IO ()

foreign import ccall "format_set_bold"
  format_set_bold :: Ptr LxwFormat_ -> IO ()

foreign import ccall "format_set_italic"
  format_set_italic :: Ptr LxwFormat_ -> IO ()

foreign import ccall "format_set_underline"
  format_set_underline :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_font_strikeout"
  format_set_font_strikeout :: Ptr LxwFormat_ -> IO ()

foreign import ccall "format_set_font_script"
  format_set_font_script :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_num_format_index"
  format_set_num_format_index :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_align"
  format_set_align :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_text_wrap"
  format_set_text_wrap :: Ptr LxwFormat_ -> IO ()

foreign import ccall "format_set_rotation"
  format_set_rotation :: Ptr LxwFormat_ -> Int16 -> IO ()

foreign import ccall "format_set_shrink"
  format_set_shrink :: Ptr LxwFormat_ -> IO ()

foreign import ccall "format_set_pattern"
  format_set_pattern :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_bg_color"
  format_set_bg_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_fg_color"
  format_set_fg_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_border"
  format_set_border :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_bottom"
  format_set_bottom :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_top"
  format_set_top :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_left"
  format_set_left :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_right"
  format_set_right :: Ptr LxwFormat_ -> Word8 -> IO ()

foreign import ccall "format_set_border_color"
  format_set_border_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_bottom_color"
  format_set_bottom_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_top_color"
  format_set_top_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_left_color"
  format_set_left_color :: Ptr LxwFormat_ -> Int32 -> IO ()

foreign import ccall "format_set_right_color"
  format_set_right_color :: Ptr LxwFormat_ -> Int32 -> IO ()
