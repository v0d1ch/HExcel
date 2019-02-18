HExcel
------

[![Hackage](https://img.shields.io/hackage/v/HExcel.svg?style=flat)](https://hackage.haskell.org/package/HExcel)


## Create Excel files with Haskell


This is a fork of [libxlsxwriter](https://github.com/HalfWayMan/libxlsxwriter)
that tries the improve on the api and provide a library for
creation of Excel files.
Underneath the hood it uses C library called [libxlsxwriter](http://libxlsxwriter.github.io/) and provides
bindings to C code to produce Excel 2007+ xlsx files.

#### Example

```
{-# LANGUAGE TypeApplications #-}
module Main where

import Control.Monad.Trans.State (execStateT)
import Control.Monad (forM_)
import Data.Time (getZonedTime)
import HExcel

main :: IO ()
main = do
  wb <- workbookNew "test.xlsx"
  let props = mkDocProperties
        { docPropertiesTitle   = "Test Workbook"
        , docPropertiesCompany = "HExcel"
        }
  workbookSetProperties wb props
  ws <- workbookAddWorksheet wb "First Sheet"
  df <- workbookAddFormat wb
  formatSetNumFormat df "mmm d yyyy hh:mm AM/PM"
  now <- getZonedTime
  -- You can create HExcelState which is convenient api for writing to cells 
  let initState = HExcelState Nothing ws 4 1 0 1 0
  _ <- flip execStateT initState $ do
         writeCell "David"
         writeCell "Dimitrije"
		 -- we can skip some rows
         skipRows 1
         writeCell "Jovana"
		 -- or skip some columns
         skipCols 1
         writeCell (zonedTimeToDateTime now)
         writeCell @Double 42.5

  -- or use functions that run in plain IO
  forM_ [5 .. 8] $ \n -> do
    writeString ws Nothing n 3 "xxx"
    writeNumber ws Nothing n 4 1234.56
    writeDateTime ws (Just df) n 5 (zonedTimeToDateTime now)
  workbookClose wb

```
