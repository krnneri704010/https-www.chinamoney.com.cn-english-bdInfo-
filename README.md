# https-www.chinamoney.com.cn-english-bdInfo-
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


def main():
    driver = webdriver.Chrome()
    driver.maximize_window()

    try:
        driver.get("https://www.chinamoney.com.cn/english/bdInfo/")
        print("页面加载成功")

        time.sleep(2)

        bond_type_select = Select(driver.find_element(By.ID, "Bond_Type_select"))
        bond_type_select.select_by_value("100001")  # Treasury Bond
        print("已选择债券类型: Treasury Bond")
        time.sleep(1)

        # 选择发行年份
        issue_year_select = Select(driver.find_element(By.ID, "Issue_Year_select"))
        issue_year_select.select_by_value("2023")  # 2023年
        print("已选择发行年份: 2023")
        time.sleep(1)

        # 点击搜索按钮
        search_button = driver.find_element(By.XPATH, '//a[@onclick="searchData()"]')
        search_button.click()
        print("点击搜索按钮，等待数据加载...")

        time.sleep(3)

        bond_data = extract_bond_data(driver)

        if bond_data:
            # 保存数据到Excel
            save_to_excel(bond_data, "bond_information_2023.xlsx")
            print(f"成功爬取 {len(bond_data)} 条债券记录")
        else:
            print("未找到债券数据，请检查选择条件或网页结构")

    except Exception as e:
        print(f"爬取过程中出现错误: {e}")
        # 截图保存用于调试
        driver.save_screenshot("error_screenshot.png")
        print("已保存错误截图: error_screenshot.png")

    finally:
        # 保持浏览器打开一段时间供查看
        input("按Enter键关闭浏览器...")
        driver.quit()


def extract_bond_data(driver):
    """提取债券信息表格数据"""
    bond_data = []

    try:
        # 等待表格出现
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )

        # 查找表格
        tables = driver.find_elements(By.TAG_NAME, "table")
        print(f"找到 {len(tables)} 个表格")

        if not tables:
            print("未找到表格元素")
            return bond_data

        # 通常数据在第一个表格中
        table = tables[0]

        # 提取表头
        headers = []
        header_rows = table.find_elements(By.TAG_NAME, "th")
        if header_rows:
            headers = [header.text.strip() for header in header_rows]
            print(f"表头: {headers}")
        else:
            # 如果没有th标签，尝试从第一行tr中提取
            first_row = table.find_elements(By.TAG_NAME, "tr")[0]
            header_cells = first_row.find_elements(By.TAG_NAME, "td")
            headers = [cell.text.strip() for cell in header_cells]
            print(f"从第一行提取的表头: {headers}")

        # 提取数据行（跳过表头行）
        rows = table.find_elements(By.TAG_NAME, "tr")
        print(f"找到 {len(rows)} 行数据")

        for i, row in enumerate(rows):
            # 跳过表头行
            if i == 0 and headers:
                continue

            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 5:  # 确保有足够的列数
                row_data = {
                    'ISIN': cells[0].text.strip() if len(cells) > 0 else '',
                    'Bond_Code': cells[1].text.strip() if len(cells) > 1 else '',
                    'Issuer': cells[2].text.strip() if len(cells) > 2 else '',
                    'Bond_Type': cells[3].text.strip() if len(cells) > 3 else '',
                    'Issue_Date': cells[4].text.strip() if len(cells) > 4 else ''
                }
                bond_data.append(row_data)
                print(f"第{i}行数据: {row_data}")

    except TimeoutException:
        print("表格加载超时，可能无数据或选择条件有误")
    except Exception as e:
        print(f"提取数据时出现错误: {e}")

    return bond_data


def save_to_excel(data, filename):
    """保存数据到Excel文件"""
    if not data:
        print("没有数据可保存")
        return

    try:
        df = pd.DataFrame(data)

        # 保存到Excel
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"数据已保存到: {filename}")

        # 打印数据预览
        print("\n数据预览:")
        print(df.head())

    except ImportError:
        # 如果openpyxl不可用，保存为CSV
        print("openpyxl未安装，将保存为CSV格式")
        csv_filename = filename.replace('.xlsx', '.csv')
        df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
        print(f"数据已保存到: {csv_filename}")


def alternative_extraction_method(driver):
    """备用的数据提取方法"""
    bond_data = []

    try:
        # 方法1: 通过特定的CSS选择器查找表格
        table_selectors = [
            'table.table',
            'table.data-table',
            'table.grid',
            'div.table-container table'
        ]

        for selector in table_selectors:
            try:
                table = driver.find_element(By.CSS_SELECTOR, selector)
                rows = table.find_elements(By.TAG_NAME, "tr")
                if len(rows) > 1:  # 至少有表头和数据行
                    print(f"通过选择器 '{selector}' 找到表格")
                    return extract_from_table_rows(rows)
            except NoSuchElementException:
                continue

        # 方法2: 查找包含债券信息的div
        div_elements = driver.find_elements(By.CSS_SELECTOR, 'div[class*="bond"], div[class*="data"]')
        for div in div_elements:
            text = div.text
            if "ISIN" in text or "Bond Code" in text:
                print("找到包含债券信息的div")
                # 进一步解析div内容

    except Exception as e:
        print(f"备用提取方法失败: {e}")

    return bond_data


def extract_from_table_rows(rows):
    """从表格行中提取数据"""
    data = []
    for i, row in enumerate(rows):
        if i == 0:  # 跳过表头
            continue
        cells = row.find_elements(By.TAG_NAME, "td")
        if cells:
            row_data = [cell.text.strip() for cell in cells]
            data.append(row_data)
            print(f"提取行 {i}: {row_data}")
    return data


if __name__ == '__main__':
    main()
/编程2
import re
from typing import List, Union


def reg_search(text: str, regex_list: List[str]) -> List[str]:
    """
    自定义正则匹配函数

    参数:
    text (str): 需要正则匹配的文本内容
    regex_list (List[str]): 正则表达式列表

    返回:
    List[str]: 匹配到的结果列表
    """
    matches = []  # 存储所有匹配结果

    # 遍历正则表达式列表中的每个模式
    for regex_pattern in regex_list:
        try:
            # 使用re.findall查找所有匹配项
            pattern_matches = re.findall(regex_pattern, text)

            # 将匹配结果添加到总列表中
            if pattern_matches:
                # 处理findall返回的不同格式（字符串列表或元组列表）
                if isinstance(pattern_matches[0], tuple):
                    # 如果是分组匹配，将元组转换为字符串
                    for match_tuple in pattern_matches:
                        # 过滤掉空分组，合并非空分组
                        valid_groups = [str(item) for item in match_tuple if item]
                        if valid_groups:
                            matches.extend(valid_groups)
                else:
                    # 如果是普通匹配，直接添加
                    matches.extend([str(match) for match in pattern_matches])

        except re.error as e:
            print(f"正则表达式错误: {regex_pattern} - 错误信息: {e}")
            continue  # 跳过无效的正则表达式

    return matches


# 测试示例
def test_reg_search():
    """测试函数"""
    # 示例文本（根据图片内容）
    text = """
    标的证券：本期发行的证券为可交换为发行人所持中国长江电力股份
    有限公司股票（股票代码：600900.SH，股票简称：长江电力）的可交换公司债券。
    换股期限：本期可交换公司债券换股期限自可交换公司债券发行结束
    之日满 12 个月后的第一个交易日起至可交换债券到期日止，即2023年 6 月 2 日至 2027 年 6 月 1 日止。
    """

    # 根据文本内容设计一些常用的正则表达式模式
    regex_list = [
        # 匹配股票代码（如：600900.SH）
        r'[0-9]{6}\.[A-Z]{2}',
        # 匹配股票简称（如：长江电力）
        r'股票简称：([\u4e00-\u9fa5]+)',
        # 匹配日期范围（如：2023年 6 月 2 日至 2027 年 6 月 1 日）
        r'\d{4}年\s*\d{1,2}月\s*\d{1,2}日至\d{4}年\s*\d{1,2}月\s*\d{1,2}日',
        # 匹配单个日期
        r'\d{4}年\s*\d{1,2}月\s*\d{1,2}日',
        # 匹配年份范围
        r'\d{4}年至\d{4}年',
        # 匹配公司名称（中文字符）
        r'[\u4e00-\u9fa5股份有限公司]+',
        # 匹配证券类型
        r'可交换公司债券|公司债券|债券',
        # 匹配数字（如：12个月）
        r'\d+个月',
    ]

    # 执行匹配
    results = reg_search(text, regex_list)

    # 打印结果
    print("原始文本:")
    print(text)
    print("\n正则表达式列表:")
    for i, pattern in enumerate(regex_list, 1):
        print(f"{i}. {pattern}")

    print("\n匹配结果:")
    for i, result in enumerate(results, 1):
        print(f"{i}. {result}")

    return results


# 增强版函数，提供更多匹配信息
def reg_search_enhanced(text: str, regex_list: List[str], return_groups: bool = False) -> List[dict]:
    """
    增强版正则匹配函数，返回更详细的信息

    参数:
    text (str): 需要正则匹配的文本内容
    regex_list (List[str]): 正则表达式列表
    return_groups (bool): 是否返回分组信息

    返回:
    List[dict]: 包含匹配详细信息的字典列表
    """
    matches_info = []

    for i, regex_pattern in enumerate(regex_list):
        try:
            # 使用re.finditer获取更详细的信息
            pattern_matches = list(re.finditer(regex_pattern, text))

            for match in pattern_matches:
                match_info = {
                    'pattern': regex_pattern,
                    'match': match.group(),
                    'start': match.start(),
                    'end': match.end(),
                    'pattern_index': i
                }

                if return_groups and match.groups():
                    match_info['groups'] = match.groups()

                matches_info.append(match_info)

        except re.error as e:
            print(f"正则表达式错误: {regex_pattern} - 错误信息: {e}")
            continue

    return matches_info


# 简单的去重版本
def reg_search_unique(text: str, regex_list: List[str]) -> List[str]:
    """
    去重版本的正则匹配函数
    """
    matches = reg_search(text, regex_list)
    # 使用集合去重，保持顺序
    seen = set()
    unique_matches = []

    for match in matches:
        if match not in seen:
            seen.add(match)
            unique_matches.append(match)

    return unique_matches


if __name__ == "__main__":
    # 运行测试
    print("=== 基本版本测试 ===")
    basic_results = test_reg_search()

    print("\n=== 增强版本测试 ===")
    enhanced_results = reg_search_enhanced(
        """
        标的证券：本期发行的证券为可交换为发行人所持中国长江电力股份
        有限公司股票（股票代码：600900.SH，股票简称：长江电力）的可交换公司债券。
        换股期限：本期可交换公司债券换股期限自可交换公司债券发行结束
        之日满 12 个月后的第一个交易日起至可交换债券到期日止，即2023年 6 月 2 日至 2027 年 6 月 1 日止。
        """,
        [r'[0-9]{6}\.[A-Z]{2}', r'股票简称：([\u4e00-\u9fa5]+)', r'\d+个月'],
        return_groups=True
    )

    for result in enhanced_results:
        print(f"模式: {result['pattern']}")
        print(f"匹配: '{result['match']}' (位置: {result['start']}-{result['end']})")
        if 'groups' in result:
            print(f"分组: {result['groups']}")
        print("-" * 50)
