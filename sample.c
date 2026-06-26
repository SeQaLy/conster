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

/* 関数別抽出 (ExtractByFunction) の動作確認用 */
int init_buffer(int mode)
{
    char buf[BUF_SIZE];          /* BUF_SIZE を使用 */
    int flags = FLAG_ALL;        /* FLAG_ALL を使用 */
    if (mode == BLUE) {          /* enum BLUE を使用 */
        flags = FLAG_A;          /* FLAG_A を使用 */
    }
    return flags & HEX_MASK;     /* HEX_MASK を使用 */
}

void draw(void)
{
    int color = WHITE;           /* enum WHITE を使用 */
    int limit = MAX_LINE;        /* MAX_LINE を使用 */
    (void)color;
    (void)limit;
}
