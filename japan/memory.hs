module Memory (
    Word32
    , allocBytes
    , coAlloc
    )

where

import Data.Word ( Word32 )
import Foreign.Ptr    (Ptr)




allocBytes :: Int -> IO (Ptr a)
allocBytes len = coAlloc (fromIntegral len)

-- | @coAlloc sz@ allocates @sz@ bytes from the COM task allocator, returning a pointer.
-- The onus is on the caller to constrain the type of that pointer to capture what the
-- allocated memory points to.
coAlloc :: Word32 -> IO (Ptr a)
coAlloc sz = allocMemory sz

--Primitives/helpers:

allocMemory :: Word32 -> IO (Ptr a)
allocMemory sz = do
  a <- primAllocMemory sz
  if a == nullPtr then
     ioError (userError "allocMemory: not enough memory")
   else
     return (castPtr a)
