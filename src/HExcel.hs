-- |
-- Module      :  HExcel
-- Maintainer  :  Sasha Bogicevic <sasa.bogicevic@pm.me>
-- Stability   :  experimental
--
-- Module exports
--
-- Example usage:
--
-- > {-# LANGUAGE TypeApplications #-}
-- > module Main where
-- >
-- > import Control.Monad.Trans.State (execStateT)
-- > import Control.Monad (forM_)
-- > import Data.Time (getZonedTime)
-- > import HExcel
-- >
-- > main :: IO ()
-- > main = do
-- >   wb <- workbookNew "test.xlsx"
-- >   let props = mkDocProperties
-- >         { docPropertiesTitle   = "Test Workbook"
-- >         , docPropertiesCompany = "HExcel"
-- >         }
-- >   workbookSetProperties wb props
-- >   ws <- workbookAddWorksheet wb "First Sheet"
-- >   df <- workbookAddFormat wb
-- >   formatSetNumFormat df "mmm d yyyy hh:mm AM/PM"
-- >   now <- getZonedTime
-- >   -- You can create HExcelState which is convenient api for writing to cells
-- >   let initState = HExcelState Nothing ws 4 1 0 1 0
-- >   _ <- flip execStateT initState $ do
-- >          writeCell "David"
-- >          writeCell "Dimitrije"
-- >          -- we can skip some rows
-- >          skipRows 1
-- >          writeCell "Jovana"
-- >          -- skip some columns
-- >          skipCols 1
-- >          writeCell (zonedTimeToDateTime now)
-- >          writeCell @Double 42.5
-- >
-- >   -- or use functions that run in plain IO
-- >   forM_ [5 .. 8] $ \n -> do
-- >     writeString ws Nothing n 3 "xxx"
-- >     writeNumber ws Nothing n 4 1234.56
-- >     writeDateTime ws (Just df) n 5 (zonedTimeToDateTime now)
-- >   workbookClose wb

module HExcel
  ( module HExcel.HExcelInternal
  )
  where

import HExcel.HExcelInternal
