module KpiStructure where

import qualified Data.Text as T 
import Data.Aeson (Value(..))

type Services = [ServStruct]


data ServStruct = ServStruct { servName :: T.Text 
                             , kpis :: [KpiStruct] 
                             } deriving (Show)

data KpiStruct = KpiStruct { kpiName :: T.Text
                           , kpiData :: Value
                           } deriving (Show)
