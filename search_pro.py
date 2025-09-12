import pymysql

conn = pymysql.connect(
    host="localhost",
    user="root",
    password="12345",
    database="user",
    port=3306,
    charset="utf8mb4"
)
select_sql = "SELECT id, chapter_no, title, question_type, options, correct_answer,short_answer FROM exam_questions WHERE title LIKE %s;"
cursor = conn.cursor()


def search_pro(title):
    cursor.execute(select_sql, "%"+title+"%")
    res = cursor.fetchall()
    print(f'共找到{len(res)}道类似题目')
    print('-' * 20)
    for i, v in enumerate(res):
        print(f'章节编号:{v[1]}')
        print(f'题目名称:{v[2]}')
        options = ' '.join(v[4].replace("\n", "").split("<|>")) if v[4] is not None else None
        print(f'选项:{options}')
        print(f'正确答案:{v[5]}')
        print(f'问题类型:{v[3]}')
        print('-' * 20)


if __name__ == '__main__':
    while True:
        title = input("请输入题目(模糊查询):")
        search_pro(title)
