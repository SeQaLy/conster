/* CConstExtractor 動作確認用サンプル */

#define MAX_LINE      128
#define HEADER_SIZE   16
#define BUF_SIZE      (MAX_LINE * 2 + HEADER_SIZE)   /* 別定数を参照 */
#define FLAG_A        (1 << 0)
#define FLAG_B        (1 << 1)
#define FLAG_ALL      (FLAG_A | FLAG_B)
#define HEX_MASK      0xFF00
#define GREETING      "hello"          // 文字列

/* 行継続つきマクロ */
#define LONG_VALUE    (MAX_LINE + \
                       HEADER_SIZE)

/* 関数形式マクロ (値は評価しない) */
#define MIN(a, b)     ((a) < (b) ? (a) : (b))

enum Color {
    RED,            /* 0 */
    GREEN,          /* 1 */
    BLUE = 5,       /* 5 */
    WHITE,          /* 6 */
    BLACK = BLUE + 10
};

const int table[BUF_SIZE];                 /* 要素数は BUF_SIZE を追って算出 */
const int matrix[3][4];                    /* 3*4 = 12 */
const int primes[] = { 2, 3, 5, 7, 11 };   /* 初期化子から 5 */
const char name[] = "abcd";                /* 4 + NUL = 5 */
const double PI = 3.14159;
static const int OFFSET = HEADER_SIZE + 1;
