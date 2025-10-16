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
