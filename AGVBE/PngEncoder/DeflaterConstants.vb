Public Class DeflaterConstants
    Public Const DEBUGGING As Boolean = False
    Public Const STORED_BLOCK As Integer = 0
    Public Const STATIC_TREES As Integer = 1
    Public Const DYN_TREES As Integer = 2
    Public Const PRESET_DICT As Integer = &H20
    Public Const DEFAULT_MEM_LEVEL As Integer = 8
    Public Const MAX_MATCH As Integer = 258
    Public Const MIN_MATCH As Integer = 3
    Public Const MAX_WBITS As Integer = 15
    Public Const WSIZE As Integer = 1 << MAX_WBITS
    Public Const WMASK As Integer = WSIZE - 1
    Public Const HASH_BITS As Integer = DEFAULT_MEM_LEVEL + 7
    Public Const HASH_SIZE As Integer = 1 << HASH_BITS
    Public Const HASH_MASK As Integer = HASH_SIZE - 1
    Public Const HASH_SHIFT As Integer = (HASH_BITS + MIN_MATCH - 1) / MIN_MATCH
    Public Const MIN_LOOKAHEAD As Integer = MAX_MATCH + MIN_MATCH + 1
    Public Const MAX_DIST As Integer = WSIZE - MIN_LOOKAHEAD
    Public Const PENDING_BUF_SIZE As Integer = 1 << (DEFAULT_MEM_LEVEL + 8)
    Public Shared MAX_BLOCK_SIZE As Integer = Math.Min(65535, PENDING_BUF_SIZE - 5)
    Public Const DEFLATE_STORED As Integer = 0
    Public Const DEFLATE_FAST As Integer = 1
    Public Const DEFLATE_SLOW As Integer = 2
    Public Shared GOOD_LENGTH As Integer() = {0, 4, 4, 4, 4, 8, 8, 8, 32, 32}
    Public Shared MAX_LAZY As Integer() = {0, 4, 5, 6, 4, 16, 16, 32, 128, 258}
    Public Shared NICE_LENGTH As Integer() = {0, 8, 16, 32, 16, 32, 128, 128, 258, 258}
    Public Shared MAX_CHAIN As Integer() = {0, 4, 8, 32, 16, 32, 128, 256, 1024, 4096}
    Public Shared COMPR_FUNC As Integer() = {0, 1, 1, 1, 1, 2, 2, 2, 2, 2}
End Class
