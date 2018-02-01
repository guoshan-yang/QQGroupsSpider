from bottle import *
import requests
from random import random
import json
from time import time, sleep
from io import BytesIO
import zipfile
import openpyxl
import unicodecsv as csv
from uuid import uuid4

attachments = {}
sourceURL = 'http://find.qq.com/index.html?version=1&im_version=5533&width=910&height=610&search_target=0'


class QQGroups:
    def __init__(self):
        self.js_ver = '10226'
        self.newSession()

    def newSession(self):
        self.sess = requests.Session()
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.59 QQ/8.9.3.21169 Safari/537.36'
        }
        self.sess.headers.update(headers)
        return

    def getQRCode(self):
        try:
            url = 'http://ui.ptlogin2.qq.com/cgi-bin/login'
            params = {
                'appid': '715030901',
                'daid': '73',
                'pt_no_auth': '1',
                's_url': sourceURL
            }
            resp = self.sess.get(url, params=params, timeout=1000)
            pattern = '//imgcache.qq.com/ptlogin/ver/(\d+)/js'
            try:
                self.js_ver = re.search(pattern, resp.text).group(1)
            except BaseException:
                pass

            print(self.js_ver)

            self.sess.headers.update({'Referer': url})
            url = 'http://ptlogin2.qq.com/ptqrshow'
            params = {
                'appid': '715030901',
                'e': '2',
                'l': 'M',
                's': '3',
                'd': '72',
                'v': '4',
                't': '%.17f' % (random()),
                'daid': '73'
            }
            resp = self.sess.get(url, params=params, timeout=1000)
            response.set_header('Content-Type', 'image/png')
            response.add_header('Cache-Control', 'no-cache, no-store')
            response.add_header('Pragma', 'no-cache')
        except BaseException:
            resp = None
        return resp

    def grLogin(self):
        login_sig = self.sess.cookies.get_dict().get('pt_login_sig', '')
        qrsig = self.sess.cookies.get_dict().get('qrsig', '')
        status = -1
        errorMsg = ''
        if all([login_sig, qrsig]):
            url = 'http://ptlogin2.qq.com/ptqrlogin'
            params = {
                'u1': sourceURL,
                'ptqrtoken': self.genqrtoken(qrsig),
                'ptredirect': '1',
                'h': '1',
                't': '1',
                'g': '1',
                'from_ui': '1',
                'ptlang': '2052',
                'action': '0-0-%d' % (time() * 1000),
                'js_ver': self.js_ver,
                'js_type': '1',
                'login_sig': login_sig,
                'pt_uistyle': '40',
                'aid': '715030901',
                'daid': '73'
            }
            try:
                resp = self.sess.get(url, params=params, timeout=1000)
                result = resp.text
                if '二维码未失效' in result:
                    status = 0
                elif '二维码认证中' in result:
                    status = 1
                elif '登录成功' in result:
                    status = 2
                elif '二维码已失效' in result:
                    status = 3
                else:
                    errorMsg = str(result.text)
            except BaseException:
                try:
                    errorMsg = str(resp.status_code)
                except BaseException:
                    pass
        loginResult = {
            'status': status,
            'time': time(),
            'errorMsg': errorMsg,
        }
        resp = json.dumps(loginResult)
        response.set_header('Content-Type', 'application/json; charset=UTF-8')
        response.add_header('Cache-Control', 'no-cache; must-revalidate')
        response.add_header('Expires', '-1')
        return resp

    def qqunSearch(self, request):
        sort = request.POST.get('sort')
        pn = int(request.POST.get('pn'))
        ft = request.POST.get('ft')
        kws = request.POST.getunicode('kws').strip()
        if not kws:
            redirect('/qqun')
            return
        kws = re.sub(r'[\r\n]', '\t', kws)
        kws = [k.strip() for k in kws.split('\t') if k.strip()]
        self.sess.headers.update({'Referer': sourceURL})
        skey = self.sess.cookies.get_dict().get('skey', '')

        try:
            buff = BytesIO()
            zip_archive = zipfile.ZipFile(buff, mode='w')
            temp = []
            for i in range(len(kws)):
                temp.append(BytesIO())
            for i, kw in enumerate(kws[:10]):
                groups = [(u'群名称', u'群号', u'群人数', u'群上限',
                           u'群主', u'地域', u'分类', u'标签', u'群简介')]
                gListRaw = []
                for page in range(0, pn):
                    # sort type: 0 deafult, 1 menber, 2 active
                    url = 'http://qun.qq.com/cgi-bin/group_search/pc_group_search'
                    data = {
                        'k': u'交友',
                        'n': '8',
                        'st': '1',
                        'iso': '1',
                        'src': '1',
                        'v': '4903',
                        'bkn': self.genbkn(skey),
                        'isRecommend': 'false',
                        'city_id': '0',
                        'from': '1',
                        'keyword': kw,
                        'sort': sort,
                        'wantnum': '24',
                        'page': page,
                        'ldw': self.genbkn(skey)
                    }
                    resp = self.sess.post(url, data=data, timeout=1000)
                    if resp.status_code != 200:
                        print('%s\n%s' % (resp.status_code, resp.text))
                    print(resp.text)
                    result = json.loads(resp.content)
                    gList = result['group_list']
                    gListRaw.extend(gList)
                    for g in gList:
                        name = self.rmWTS(g['name'])
                        code = g['code']
                        member_num = g['member_num']
                        max_member_num = g['max_member_num']
                        owner_uin = g['owner_uin']
                        qaddr = ' '.join(g['qaddr'])
                        try:
                            gcate = ' | '.join(g['gcate'])
                        except BaseException:
                            gcate = ''
                        try:
                            _labels = [l.get('label', '') for l in g['labels']]
                            labels = self.rmWTS(' | '.join(_labels))
                        except BaseException:
                            labels = ''
                        memo = self.rmWTS(g['memo'])
                        gMeta = (name, code, member_num, max_member_num,
                                 owner_uin, qaddr, gcate, labels, memo)
                        groups.append(gMeta)
                    if len(gList) == 1:
                        break
                    sleep(2.5)
                if ft == 'xls':
                    wb = openpyxl.Workbook()
                    sheet = wb.active
                    sheet.title = 'QQ群数据表'
                    for k in range(0, len(groups)):
                        for j in range(0, len(groups[k])):
                            sheet.cell(row=k + 1, column=j + 1,
                                       value=str(groups[k][j]))
                    wb.save(temp[i])
                elif ft == 'csv':
                    writer = csv.writer(
                        temp[i], dialect='excel', encoding='utf-8')
                    writer.writerows(groups)
                elif ft == 'json':
                    temp[i].write((json.dumps(gListRaw).encode("utf-8")))
                    # json.dump(gListRaw, temp[i], indent=4, sort_keys=True)
            for i in range(len(kws)):
                zip_archive.writestr(kws[i] + '.' + ft, temp[i].getvalue())
            zip_archive.close()
            resultId = uuid4().hex
            attachments.update({resultId: buff})
            response.set_header('Content-Type', 'text/html; charset=UTF-8')
            response.add_header('Cache-Control', 'no-cache; must-revalidate')
            response.add_header('Expires', '-1')
            return resultId
        except Exception as e:
            print(e)
            abort(500,)

    def genqrtoken(self, qrsig):
        e = 0
        for i in range(0, len(qrsig)):
            e += (e << 5) + ord(qrsig[i])
        qrtoken = (e & 2147483647)
        return str(qrtoken)

    def genbkn(self, skey):
        b = 5381
        for i in range(0, len(skey)):
            b += (b << 5) + ord(skey[i])
        bkn = (b & 2147483647)
        return str(bkn)

    def rmWTS(self, content):
        pattern = r'\[em\]e\d{4}\[/em\]|&nbsp;|<br>|[\r\n\t]'
        content = re.sub(pattern, ' ', content)
        content = content.replace('&amp;', '&').strip()
        return content


bottle = Bottle()
q = QQGroups()


@bottle.route('/')
def home():
    redirect('/qqun')


@bottle.route('/getqrcode')
def getQRCode():
    return q.getQRCode()


@bottle.route('/qrlogin')
def grLogin():
    return q.grLogin()


@bottle.route('/static/<path:path>')
def server_static(path):
    return static_file(path, root='static')


@bottle.route('/qqun', method='ANY')
def qqun():
    if request.method == 'GET':
        response.set_header('Content-Type', 'text/html; charset=UTF-8')
        response.add_header('Cache-Control', 'no-cache')
        return template('qqun')
    elif request.method == 'POST':
        return q.qqunSearch(request)


@bottle.route('/download')
def download():
    resultId = request.query.rid or ''
    f = attachments.get(resultId, '')
    if f:
        response.set_header('Content-Type', 'application/zip')
        response.add_header('Content-Disposition',
                            'attachment; filename="results.zip"')
        return f.getvalue()
    else:
        abort(404)


if __name__ == '__main__':
    run(bottle, host='localhost', port=8080, debug=True, reloader=True)
