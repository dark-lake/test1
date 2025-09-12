import copy
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pymysql
from data_struct import my_item, tm_type

'''
该脚本的目的在于爬取答题完毕后的题目,并将其中的正确题目导入到文件中
1. 手动输入答题完毕后的地址,启动脚本
2. 抓取页面,并将题目全部获取
3. 将其中重复的题目过滤
4. 将正确的题目标记出来
'''

conn = pymysql.connect(
    host="localhost",
    user="root",
    password="12345",
    database="user",
    port=3306,
    charset="utf8mb4"
)
insert_sql = "INSERT INTO exam_questions(chapter_no, title, question_type, options, correct_answer, short_answer) VALUES (%s, %s, %s, %s, %s, %s);"
select_sql = "SELECT id, chapter_no, title, question_type, options, correct_answer,short_answer FROM exam_questions WHERE title LIKE %s;"
update_sql = "UPDATE exam_questions SET chapter_no = %s,title = %s,question_type = %s,options = %s,correct_answer = %s,short_answer = %s WHERE id = %s;"
cursor = conn.cursor()


def start(url):
    # 设置不自动关闭
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # 让浏览器不自动关闭
    driver = webdriver.Chrome(executable_path="E:\chromedriver-win64\chromedriver-win64\chromedriver.exe",
                              options=options)
    driver.maximize_window()
    driver.get(url)
    print(driver.page_source)

    # 开始登录
    # 找到企业微信文字
    qywx_text = driver.find_element(By.XPATH, '//*[@id="login"]/div/div[1]/div[8]/ul/li[2]/a/span')
    # 点击后弹出二维码
    qywx_text.click()
    # 扫码登录后,可以直接跳转到指定课程考试即可
    driver.get(
        'https://lexiangla.com/exams/2ccbb7f2a19311efa9853aace62fdd0e?company_from=ec8b337e5e4c11edb2296a452cf85f24')

    # 等待某个元素出现
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="exam-container"]//button[contains(text(), "查看结果")]'))
    )

    print(element.text)

    # 获取查看结果按钮,并点击
    print(driver.page_source)
    driver.find_element(By.XPATH, '//*[@id="exam-container"]//button[contains(text(), "查看结果")]').click()

    # 获取所有考试结果的url
    result_list = driver.find_elements(By.XPATH, '//*[@id="select-exam-result-modal"]/div/div/div//a')
    print(f'获取到考试结果次数为:{len(result_list)}')
    if len(result_list) == 0:
        print("未获取到考试结果, 结束执行")
        return
    res_list = [i.get_attribute('href') for i in result_list]
    print(f'考试href为:{res_list}')

    # 跳转考试结果
    for i, v in enumerate(res_list):
        print(f'开始处理第{i + 1}次考试')
        driver.get(v)
        # 开始抓取题目数据
        # 等待某个元素出现
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//div[@class="question-item"]'))
        )

        print(element.text)

        # 获取章节标题
        chapter_title = None
        try:
            chapter_title = driver.find_element(By.XPATH,
                                                '//*[@id="app-vue"]//div[contains(@class, "exam-title")]').text
        except Exception as e:
            print(f'未找到章节标题:{chapter_title}, 报错信息:{e}')
            return

        # 获取所有题目
        all_ques = driver.find_elements(By.XPATH, '//div[@class="question-item"]')
        if len(all_ques) == 0:
            print(f'未获取到题目在{i + 1}次考试')
            return

        # 获取不同题目类型编号
        ques_range = driver.find_elements(By.XPATH, '//ul[@class="answer-thumbnail clearfix"]/li')
        if len(ques_range) == 0:
            print(f'未获取到题号区间在{i + 1}次考试')
            return

        if len(all_ques) != len(ques_range):
            print(f'题目数与题号区间不匹配{len(all_ques)}:{len(ques_range)}在{i + 1}次考试')
            return

        # 找出第一个类型题目一共有多少道,已知每一个问题的第一个都有一个question-block快,所以可以判定第一个就是他的题型行
        # 先拿到所有的个数,然后找到每个block块的位置,然后根据位置和题号来切分对象类型
        # type_index 记录不同题目类型的开始位置
        my_tm_type = copy.deepcopy(tm_type)
        for p, u in enumerate(ques_range):
            # 用于记录当前索引是否是chapter位置
            temp = None
            try:
                temp = u.find_element(By.XPATH, 'div[@class="question-block"]').text
            except:
                pass
            if temp is None:
                continue
            if "单选" in temp.strip():
                my_tm_type['dan_xuan'] = p
            elif "多选" in temp.strip():
                my_tm_type['duo_xuan'] = p
            elif "判断" in temp.strip():
                my_tm_type['pan_duan'] = p

        print(f'获取题目类型编号成功:{my_tm_type}')

        # 开始保存题目
        # 获取不同类型题目的开始位置
        my_tm_keys = list(my_tm_type.keys())
        my_tm_values = list(my_tm_type.values())
        # 记录题目类型的中文名称
        question_type = None
        # 记录题型的开始索引
        start_index = None
        for k, j in enumerate(all_ques):
            item = copy.deepcopy(my_item)
            # 如果题目的索引在不同题目的开始位置中,那也就意味着这道题是
            if k in my_tm_values:
                # 获取当前题目类型中文名称
                question_type = my_tm_keys[my_tm_values.index(k)]
                start_index = k
            # 开始填充item对象
            title = ''.join(j.find_element(By.XPATH, 'div/div[@class="title"]').text.strip().split(". ")[1:])
            # 获取所有选项的拼接字符串
            if question_type != 'pan_duan':
                options = '<|>'.join([t.text.strip() for t in j.find_elements(By.XPATH, 'ul/li')])
            else:
                options = None
            # 获取到两个结果span,第一个span是你的答案,第二个span是判断你是答案是否正确
            correct_answers_span = j.find_elements(By.XPATH,
                                                  'div[contains(@class, "mt-m")]/div[contains(@class, "pl-s")]/span')
            if len(correct_answers_span) < 2:
                print(f'当前题目的获取结果span异常: {title, options},请手动处理')

            # 获取正确答案,最后用一个d来判断是对的,用x判断是错的
            your_answer = correct_answers_span[0].text.strip()
            d_or_x = 'd' if "icon-correct" in correct_answers_span[1].get_attribute("class").split(" ") else 'x'

            item['chapter_no'] = chapter_title
            item['title'] = title
            item['question_type'] = question_type
            item['options'] = options
            item['correct_answer'] = your_answer+d_or_x
            # 简单题目前还没见过

            # 先从数据库获取该题
            cursor.execute(select_sql, (item['title'],))
            rows = cursor.fetchone()
            if rows is None:
                cursor.execute(insert_sql, tuple(list(item.values())))
            else:
                # 如果数据库中的答案正确直接跳过
                if rows[5][-1] == 'd':
                    continue
                # 不正确则通过id更新数据库
                elif d_or_x == 'd':
                    temp = list(item.values())
                    # 设置where的id
                    temp.append(rows[0])
                    cursor.execute(update_sql, tuple(temp))

            print(f'{item}')
        # 跑完一个网页提交一次
        conn.commit()



if __name__ == '__main__':
    url = 'https://lexiangla.com/login?use_lexiang=1'
    start(url)
