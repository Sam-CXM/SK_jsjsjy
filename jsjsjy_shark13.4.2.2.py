from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium import webdriver
from ddddocr import DdddOcr
from time import sleep, strftime, localtime, gmtime, time
from os import makedirs, path, listdir, remove
from win32com.client import Dispatch
from wget import download
from zipfile import ZipFile
from re import findall


def main():
    """ 主函数 """
    version_code = '2040004'
    head = "开发作者：晨小明\n软件版本：v2.4\n程序版本：v13.4.4\n版本编码：" + \
        version_code + "\n维护日期：2023/04/21\n更新日志：\n\t【优化】其他bug修复\n"
    log_file('version_code: ' + version_code)
    print(head)

    sleep(0.5)
    print("Tips: 欢迎使用本程序，请保持网络流畅，全程请不要操作自动打开的窗口，否则可能出现错误。")
    print("      如果出现错误，导致异常退出，请重新运行本程序！\n")


def check_fix_sys():
    """ 检查运行环境总函数 """
    def down_driver(version1, file_path2):
        """   下载webdriver软件   """
        print("webdriver文件错误！开始修复...\r", end="")
        log_file('tip: ' + 'webdriver文件错误！开始修复')

        down_path = "https://msedgedriver.azureedge.net/" + \
                    version1 + "/edgedriver_win64.zip"
        do = download(down_path, 'D:\\')
        # print(do)
        file_path = 'D:\\edgedriver_win64.zip'
        z = ZipFile(file_path, 'r')
        z.extract('msedgedriver.exe', 'D:\\')
        z.close()
        remove(file_path)
        log_file('delete: ' + file_path)
        print("\n下载webdriver软件完成，文件路径：D:\msedgedriver.exe。\n由于系统有权限，请手动剪切到 C:\Program Files (x86)\Microsoft\Edge\Application\ 文件夹下。")
        input("按Enter键确认完成操作")
        log_file('tip: ' + '确认完成操作')
        a = 0
        while path.isfile(file_path2) == False:
            print(" ··>错误<··  操作失败！请重新操作")
            log_file('error: ' + '操作失败！重新操作')
            input("按Enter键确认完成操作")
            log_file('tip: ' + '确认完成操作')
            if a > 1:
                if path.isfile(file_path2) == True:
                    break
                else:
                    print(" ··>错误<··  操作错误次数较多，等待重新下载")
                    log_file('error: ' + '操作错误次数较多，重新下载')
                    down_driver(version1, file_path2)
            a += 1
        if path.isfile("D:\msedgedriver.exe") == True:
            remove("D:\msedgedriver.exe")
            log_file('delete: ' + 'D:\msedgedriver.exe')
        else:
            pass
        print(" ··>提示<··  环境修复完成！")
        log_file('tip: ' + '环境修复完成')

    def get_version_number():
        ''' 获取文件版本信息，这个兼容性强 '''
        sleep(0.5)
        print("正在检测系统环境...\r", end="")
        file_path1 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        file_path2 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe"
        if path.isfile(file_path1) == False:
            log_file('error: ' + 'Edge浏览器未下载或启动路径不正确！')
            log_file('return: ' + 'False')
            return False
        else:
            information_parser = Dispatch("Scripting.FileSystemObject")
            version1 = information_parser.GetFileVersion(file_path1)
            if path.isfile(file_path2) == True:
                version2 = information_parser.GetFileVersion(file_path2)
                # print(version1[:-3], version2[:-3])
                if version1[:-3] == version2[:-3]:
                    print(" ··>提示<··  环境正常！")
                    log_file('tip: ' + '环境正常！')
                else:
                    down_driver(version1, file_path2)
            else:
                down_driver(version1, file_path2)

    return get_version_number()


def start():
    """ 窗口初始化 """

    driver.get("https://www.jste.net.cn/uids/index.jsp")
    title = driver.title

    if title != "江苏教师管理":
        log_file('tip: ' + '网站维护中')
        end(" ··>提示<··  网站维护中，已关闭页面！")
    else:
        print(f"已打开 {title}\n")
        log_file('tip: ' + f"已打开 {title}")

        # 设置窗口高度和宽度
        driver.maximize_window()
        size = driver.get_window_size()
        width = size['width']
        height = size['height']
        driver.set_window_size(720, height / 1.1)

        # 设置窗口位置
        driver.set_window_position(0, 0)

        # 初始化验证码图片存储位置
        tempCodePic = 'D:\codePic.png'
        ocr = DdddOcr()
        print("---------------开始登录---------------")
        log_file('tip: ' + '开始登录')
        adminPwd(ocr, tempCodePic, admin='', password='')
        sleep(2)
        try:
            driver.find_element(
                By.CSS_SELECTOR, '#mirrorContainer > div.choosedSite > input').click()
        except:
            pass

        bestudy = getBestudyTime()
        if bestudy == True:
            end("\n ··>提示<··  所有视频播放完成，已退出页面！")
            log_file('tip: ' + '所有视频播放完成')
        else:
            isPlay()
            isPlay()


def adminPwd(ocr, tempCodePic, admin, password):
    """ 登录主函数 """
    if admin == "":
        ad_input = input("请顺序输入账号和密码（账号在前，密码在后），用空格隔开：")
        log_file('tip: ' + '输入账号和密码')

        if ad_input == "":
            print(" ··>错误<··  您的输入为空，请重新输入")
            log_file('tip: ' + '输入为空，重新输入')
            adminPwd(ocr, tempCodePic, admin='', password='')
        else:
            ad_input = ad_input.split(" ")
            if len(ad_input) != 2:
                print(" ··>错误<··  输入内容项数量不正确，请检查后重试")
                log_file('error: ' + '输入内容项数量不正确，重新输入')
                adminPwd(ocr, tempCodePic, admin='', password='')
            else:
                if ad_input[0] == "":
                    print(" ··>错误<··  您输入的账号为空，请重新输入")
                    log_file('error: ' + '输入的账号为空，重新输入')
                    adminPwd(ocr, tempCodePic, admin='', password='')
                elif ad_input[1] == "":
                    print(" ··>错误<··  您输入的密码为空，请重新输入")
                    log_file('error: ' + '输入的密码为空，重新输入')
                    adminPwd(ocr, tempCodePic, admin='', password='')
                else:
                    admin = ad_input[0]
                    password = ad_input[1]
    else:
        admin = admin
        password = password
    log_file('login_username: ' + admin)
    elem = driver.find_element(By.ID, "loginName")
    elem.clear()  # 清空上述的文本内容
    elem.send_keys(admin)
    elem = driver.find_element(By.ID, "pwd")
    elem.clear()
    elem.send_keys(password)
    elem = driver.find_element(By.ID, "randomCode")
    elem.clear()
    elem.send_keys("")
    codePic = driver.find_element(By.ID, 'imgCode')
    sleep(2)
    codePic.screenshot(tempCodePic)
    sleep(1)
    seeCode(ocr, elem, tempCodePic, admin, password)    # 调用识别验证码
    sleep(0.5)

    # 点击登录
    driver.find_element(
        By.XPATH, '/html/body/table[1]/tbody/tr[4]/td/form/input[4]').click()
    log_file('tip: ' + '点击登录')
    sleep(0.5)

    isLogin(tempCodePic, ocr, admin, password)
    sleep(2)

    # 打开课程信息
    driver.switch_to.frame(2)
    driver.find_element(
        By.XPATH, '//*[@id="competition"]/ul[3]/li[1]/a').click()
    log_file('tip: ' + '点击课程信息')
    sleep(1)

    handles()

    error = " ··>警告<··  网站错误，请退出！"
    title = driver.title
    if title == "Error":
        end(error)
        log_file('warning: ' + '网站错误')


def seeCode(ocr, elem, tempCodePic, admin, password):
    """ 验证码识别函数 """
    with open(tempCodePic, 'rb') as f:
        img_bytes = f.read()
        rdCode = ocr.classification(img_bytes)
    # print(rdCode)
    if rdCode != "":
        elem.send_keys(rdCode)
        log_file('rdCode: ' + rdCode)
    else:
        print(" ··>错误<··  由于网络原因验证码未显示，正在重新登录")
        log_file('error: ' + '验证码未显示')
        driver.refresh()
        adminPwd(ocr, tempCodePic, admin=admin, password=password)


def isLogin(tempCodePic, ocr, admin, password):
    """ 判断是否登陆成功 """
    sleep(1)
    if tempCodePic:
        remove(tempCodePic)
        log_file('delete: ' + tempCodePic)
        # print("已删除临时验证码图片")
    print("正在检测登录状态，请稍候...\r", end="")
    log_file('tip: ' + '检测登录状态')

    title = driver.title
    if title == '江苏省中小学教师(校长)培训学时认定和管理系统':
        print(' ··>提示<··  登录成功！ ' + strftime("%Y-%m-%d %H:%M:%S", localtime()))
        log_file('tip: ' + '登录成功！')
        sleep(0.1)
    else:
        errmesg = driver.find_element(
            By.XPATH, "/html/body/table[1]/tbody/tr[4]/td/ul/li/span").text
        if errmesg == "无此登录帐号":
            print(" ··>错误<··  无此登录账号，请重新输入！")
            log_file('error: ' + '无此登录账号，重新输入！')
            adminPwd(ocr, tempCodePic, admin='', password='')
        elif errmesg == "错误的登录密码":
            print(" ··>错误<··  登录密码错误，请重新输入！")
            log_file('error: ' + '登录密码错误，重新输入！')
            adminPwd(ocr, tempCodePic, admin='', password='')
        elif errmesg == "错误的验证码":
            print(" ··>错误<··  验证码错误，正在重新登录！")
            log_file('error: ' + '验证码错误，正在重新登录！')
            adminPwd(ocr, tempCodePic, admin=admin, password=password)
        elif errmesg == "您的帐号已经连续4次输错密码，再输错一次帐号将被锁定1440分钟，建议您立即联系上级管理员重置您的密码！":
            print(" ··>警告<··  您的帐号已经连续4次输错密码，再输错一次帐号将被锁定1440分钟，建议您立即联系上级管理员重置您的密码！")
            log_file('warning: ' + '连续4次输错密码')
            adminPwd(ocr, tempCodePic, admin='', password='')
        else:
            print(" ··>错误<··  未知错误，请重新输入！")
            log_file('error: ' + '未知错误，重新输入！')
            adminPwd(ocr, tempCodePic, admin='', password='')


def handles():
    """ 切换标签页 """
    handles = driver.window_handles
    driver.switch_to.window(handles[-1])
    log_file('tip: ' + '切换标签页')
    sleep(2)


def getDONGCourseCount():
    """ 动态获取课程数量函数 """
    course_count = driver.find_element(
        By.XPATH, '/html/body/div[5]/div[2]/div')
    course_count = course_count.find_elements(By.TAG_NAME, 'a')
    course_count = len(course_count) + 1
    log_file('tip: ' + '获取课程数量: ' + str(course_count))
    return course_count


def isPlay():
    """ 寻找未播放课程 """
    log_file('tip: ' + '寻找未播放课程')
    course_count = getDONGCourseCount()

    for e in range(1, course_count):
        # print("正在检测已学课程，请稍候...\r", end="")
        if e == 1:
            log_file('tip: ' + '检测已学课程')
        isstudy = driver.find_element(
            By.XPATH, '/html/body/div[5]/div[2]/div/a[' + str(e) + ']/div/div[2]/div/span').text

        isstudy = findall("\d+", isstudy)

        if int(isstudy[1]) >= int(isstudy[0]):
            subjec = driver.find_element(
                By.XPATH, '/html/body/div[5]/div[2]/div/a[' + str(e) + ']/div/div[2]/p/small').text
            # print(f"已学 {e}--{subjec}")

        else:
            # print(" ··>提示<··  检测已学课程完成！\r", end="")
            print("\n---------------开始播放---------------")
            if e == 1:
                log_file('tip: ' + '开始播放')

            play(course_count, e)

            bestudy = getBestudyTime()
            if bestudy == True:
                end("\n ··>提示<··  所有视频播放完成，已退出页面！")
                log_file('tip: ' + '所有视频播放完成')
                break
        sleep(0.1)


def getBestudyTime():
    """ 获取时长函数 """
    be_study_time = driver.find_element(
        By.XPATH, '/html/body/div[3]/div/span[1]/span[1]').text
    be_study_time = findall("\d+", be_study_time)
    print(f" ··>提示<··  您已完成了 {be_study_time[1]} / {be_study_time[0]} 分钟。")
    log_file(f'tip: 已完成了 {be_study_time[1]} / {be_study_time[0]} 分钟')
    if be_study_time[1] >= be_study_time[0]:
        return True
    else:
        return False


def play(course_count, e):
    """ 检测当前视频是否播放完成，如果播放完，点击确定进入下一课程 """
    # 点击播放
    driver.find_element(
        By.XPATH, '/html/body/div[5]/div[2]/div/a[' + str(e) + ']/div/div[1]').click()
    log_file('tip: ' + '点击播放')

    # 单个课程的总时长
    course_cnt_time = driver.find_element(
        By.XPATH, '/html/body/div[3]/div/span[1]/span[2]').text
    course_cnt_time = findall("\d+", course_cnt_time)

    # 获取视频标题
    sub_play_name = driver.find_element(
        By.XPATH, '/html/body/div[1]/div/ol/li[2]').text
    if len(sub_play_name) > 20:
        sub_pl_na = sub_play_name[:20] + '...'
    else:
        sub_pl_na = sub_play_name
    course_count -= 1
    print(
        f" ··>提示<··  正在播放：{e} / {course_count} -- {sub_pl_na}，时长：{course_cnt_time[0]}分钟")
    log_file(
        'tip: ' + f'播放：{e} / {course_count} -- {sub_pl_na}，时长：{course_cnt_time[0]}分钟')

    # 转换成时间
    sub_cnt_time = int(course_cnt_time[2]) * 60
    str_time = strftime("%H:%M:%S", gmtime(int(course_cnt_time[0]) * 60))

    log_file('tip: ' + '开始播放')

    if int(course_cnt_time[2]) > 0:
        while True:
            sub_cnt_time += 1
            be_time = strftime("%H:%M:%S", gmtime(sub_cnt_time))
            sleep(1)
            print(f"{be_time} / {str_time}\r", end="")
            if sub_cnt_time == int(course_cnt_time[0]) * 60 + 60:
                break
    else:
        t = 0
        while True:
            t += 1
            be_time = strftime("%H:%M:%S", gmtime(t))
            sleep(1)
            print(f"{be_time} / {str_time}\r", end="")
            if t == int(course_cnt_time[0]) * 60 + 60:
                break
    log_file('tip: ' + '播放完成')

    driver.back()
    driver.refresh()


def end(msg):
    """ 结束（退出）函数 """
    print(msg)


def log_file(log_info):
    """ 日志生成函数 """
    path0 = r'./log'
    if path.exists(path0) == False:
        makedirs(path0)
    ct = time()
    log_filename = strftime("%Y-%m-%d", localtime())
    data_head = strftime("%Y-%m-%d %H:%M:%S", localtime(ct))
    data_secs = (ct - int(ct)) * 1000
    time_stamp = "%s.%03d" % (data_head, data_secs)
    with open("./log/" + log_filename, "a+", encoding='utf-8') as f:
        f.write("[ " + time_stamp + " ]  " + log_info + "\n")


if __name__ == "__main__":
    log_file('\nStart Run')
    main()
    tr = check_fix_sys()
    if tr == False:
        sleep(0.5)
        print(" ··>错误<··  Edge浏览器未下载，请下载后重试！")
        input("按任意键退出程序！")
        log_file('tip: ' + '退出程序')
        print("已退出程序")
    else:
        sleep(0.5)
        print("\n--------------------开始打开网页--------------------")
        log_file('tip: ' + '打开网页')
        edge_options = Options()
        edge_options.add_experimental_option(
            'excludeSwitches', ['enable-logging'])     # 隐藏不重要错误
        edge_options.add_argument("--mute-audio")  # 静音播放
        with webdriver.Edge(
                "C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe", options=edge_options) as driver:
            driver.implicitly_wait(30)      # 隐式等待

            start()

        input("按任意键退出程序！")
        log_file('tip: ' + '退出程序')
        print("已退出程序")
