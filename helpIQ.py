# coding:utf-8
import os.path
import pdb
import xml.dom.minidom
import sys
import shutil
import io
import json
import re
import time
import win32com.client
import threading
import pythoncom
import locale
pythoncom.CoInitialize()

EXCLUDELIST = [r'OpManager_REST_API.html', r'(?i)[^"]*?xml(_\d+)?.html', r'(?i)[^"]*?xsd(_\d+)?.html', r'[^"]*_Example.html', r'Release_Notes.html']  # files that would be excluded in word(docx) file.
LIMITPICNO = 254  # If you want to save a file with too many pictures, Word may take too long time and crashed in some case. (in my case: apm - Exchange_Server_Monitoring.html -741 pictures, Word crashed.)


def myCopyTree(soure, des):
    for current, dirs, files in os.walk(soure):
        desCurrent = current.replace(soure, des, 1)
        for d in dirs:
            desDirPaht = os.path.join(desCurrent, d)
            if os.path.exists(desDirPaht):
                pass
            else:
                os.makedirs(desDirPaht)
        for f in files:
            sorFilePath = os.path.join(current, f)
            desFilePath = os.path.join(desCurrent, f)
            try:
                shutil.copy2(sorFilePath, desFilePath)
            except Exception, e:
                print '[!] Failed to copy file: {} ||| Error: {}'.format(sorFilePath, e)


def xml2JsFile(path, xmlpath):
    '''
    Parse help.xml file
    Write each page by it order to js file in list format, which will used by help template's navigator
    return pageDict for afterward usage.
    '''
    try:
        DomTree = xml.dom.minidom.parse(xmlpath)
    except xml.parsers.expat.ExpatError as e:
        print '[!] help.xml file ERROR: {}'.format(e)
        sys.exit(1024)
    collection = DomTree.documentElement
    pages = collection.getElementsByTagName('page')
    folders = collection.getElementsByTagName('folder')
    orders = collection.getElementsByTagName('order')

    pageDict = {}
    folderDict = {}
    nodeList = []
    ROOTNODE = {'node_id': '0'}
    finallList = []

    for page in pages:
        topic_status = page.getElementsByTagName('topic_status')[0].childNodes[0].data
        topic_id = page.getElementsByTagName('topic_id')[0].childNodes[0].data
        toc_name = page.getElementsByTagName('toc_name')[0].childNodes[0].data
        file_name = page.getElementsByTagName('file_name')[0].childNodes[0].data
        pageDict[topic_id] = {'toc_name': toc_name, 'file_name': file_name, 'topic_status': topic_status}
    for folder in folders:
        folder_id = folder.getElementsByTagName('folder_id')[0].childNodes[0].data
        folder_name = folder.getElementsByTagName('folder_name')[0].childNodes[0].data
        try:
            link_url = folder.getElementsByTagName('link_url')[0].childNodes[0].data
        except IndexError:
            link_url = ''
        folderDict[folder_id] = {'link_url': link_url, 'folder_name': folder_name}
    for order in orders:
        node_id = order.getElementsByTagName('node_id')[0].childNodes[0].data
        node_type = order.getElementsByTagName('node_type')[0].childNodes[0].data
        node_order = order.getElementsByTagName('node_order')[0].childNodes[0].data
        parent_id = order.getElementsByTagName('parent_id')[0].childNodes[0].data
        parent = parent_id.split(';')   # <parent_id>t;1804110</parent_id>
        parent.reverse()
        parent_id = parent[0]
        nodeList.append({'node_order': node_order, 'node_id': node_id, 'parent_id': parent_id, 'node_type': node_type})

    def walkNode(node, upperList):
        subs = list(filter(lambda x: x['parent_id'] == node['node_id'], nodeList))
        sorted_subs = sorted(subs, key=lambda x: int(x['node_order']))
        for sub_node in sorted_subs:
            target = u'MAIN'
            if sub_node['node_type'] == 'T':
                thisnode = pageDict.get(sub_node['node_id'])
                if thisnode.get('topic_status') == '0':
                    continue  # hidden page
                name = thisnode.get('toc_name').strip()
                link = thisnode.get('file_name')

            else:
                thisnode = folderDict.get(sub_node['node_id'])
                name = thisnode.get('folder_name').strip()
                link = thisnode.get('link_url')
                # I do not know how to check if a folder is hidden or not
            if re.match(r'ID\d*', link):  # <link_url>ID182360</link_url> some page list in navigator more than twice?
                linkid = link[2:]
                if linkid in pageDict:
                    link = pageDict.get(linkid)['file_name']
                elif linkid in folderDict:
                    link = folderDict.get(linkid)['link_url']
            if re.match(r'(?i)https?.*?\$hqt_blank', link):
                link = link.split('$', 1)[0]
                target = u'_BLANK'
            if link == u'.html': # this weird link come some time...
                link = u''
            if not name:  # named in blank space
                continue
            if not link.startswith(('http://', 'https://')) and not os.path.exists(os.path.join(path, link)):
                print u'[!] File does not exist: {}'.format(os.path.join(path, link))
                continue
            thisNode = [name, link, target]
            upperList.append(thisNode)
            walkNode(sub_node, thisNode)
    walkNode(ROOTNODE, finallList)
    s = 'var TREE_NODES =' + json.dumps(finallList, encoding='utf-8')
    s = s.decode('unicode-escape')
    myCopyTree('user-guide-template', path)
    with io.open(os.path.join(path, 'script/tree_nodes.js'), 'w', encoding='utf-8') as jsfile:
        jsfile.write(s)
        print u'[+] tree_nodes.js written.'
    return pageDict


def addHelpTemplate(path):
    if os.path.exists(os.path.join(path, 'index.html')):
        print '[!] Seems the User Guide Template has been applied...If not, unzip a fresh backed site and then try again.'
        return
    xmlFileName = os.path.join(path, 'help.xml')
    if not os.path.exists(xmlFileName):
        print '[!] Can not find help.xml file. Make sure the folder is a backuped Help IQ site.'
        sys.exit(1024)
    pages = xml2JsFile(path, xmlFileName)
    cleanHtmlFile(pages, path)
    inertCustomeCSS(path)
    changeIntroFile(path)
    print '[*] The User Guide Template applied.'


def cleanHtmlFile(pages, path):
    '''
    In exported page, the a tag look like: <a hqid="182387" href="#">text</a>
    Usage this function to find the file_name of this hqid, and then modify th href attritue with found file name.
    Another action in this funciton: img tag, replace internet address with local address.
    '''
    for curdir, folders, files in os.walk(path):
        linkPattern = re.compile(r'(?is)<a[^.]*?hqid="(\d*?)"[^>]*?href="#"[^>]*?>(.*?)</a>')

        def linkRepl(m):
            '''
            If in helpIQ we linked to an unexist page, it would not to find file name of such <a>'s hqid.
            Use this complex function to remove <a> tag if it is a break one. Otherwith, to repleace # with file name.
            '''
            matchedStr = m.group(0)
            pageId = m.group(1)
            pageNode = pages.get(pageId)
            if pageNode:
                fileName = pageNode['file_name']
                return matchedStr.replace(u'"#"', u'"{}"'.format(fileName))
            else:
                return m.group(2)

        imgPattern = re.compile(r'(?is)(<img[^>]*?src=")([^"]*?)"')

        def imgRepl(m):
            '''
            Some image file does not include in exported files.
            Use this complex function to check if a image file exist locally, otehrwise, use the internet address.
            '''
            fileName = m.group(2)
            names = fileName.split('images/', 1)
            names.reverse()
            localName = 'images/' + names[0].split('?', 1)[0]
            clearFileNamePath = os.path.join(path, localName)
            if os.path.exists(clearFileNamePath):
                return m.group(1) + localName + '"'
            else:
                return m.group(0)

        for f in files:
            if os.path.splitext(f)[-1] == '.html':
                filePath = os.path.join(curdir, f)
                with io.open(filePath, 'r+', encoding='utf-8') as fh:
                    try:
                        content = fh.read()
                    except Exception:
                        continue
                    content = linkPattern.sub(linkRepl, content)
                    content = imgPattern.sub(imgRepl, content)
                    fh.seek(0)
                    fh.truncate()
                    fh.write(content)
    print u'[+] html files cleaned.'


def inertCustomeCSS(path):
    customCSS = ''
    with open(os.path.join(path, 'custom.css'), 'rb') as fh:
        customCSS = fh.read()
    if customCSS:
        with open(os.path.join(path, 'style.css'), 'ab') as fh:
            fh.write(customCSS)
    print u'[+] custom css appended.'


def changeIntroFile(path):
    nodejsFile = os.path.join(path, 'script/tree_nodes.js')
    htmlFile = ''
    if os.path.exists(nodejsFile):
        with io.open(nodejsFile, 'r', encoding='utf-8') as fh:
            content = fh.read()
            html = re.search(r'(?is)"([^"]*?\.html)"', content)
            if html:
                htmlFile = os.path.join(path, html.group(1))
    if htmlFile:
        shutil.copy2(htmlFile, os.path.join(path, 'intro.html'))
        print u'[+] {} copied as intro.html, the default opened page.'.format(htmlFile)


def generateDocx(path):
    word = win32com.client.DispatchEx('Word.Application')
    # word.Visible = 1
    word.NormalTemplate.Saved = 1  # tofix : prompted to save normal.dot when quick word.
    nodejsFile = os.path.join(path, 'script/tree_nodes.js')
    htmls = []
    aPattern = re.compile(r'(?is)<a(?![^>]*https?).*?>(.*?)</a>')
    blankLinePattern = re.compile(r'(?i)<(p|a|h\d) ?[^>]*?>\s*?(&nbsp;)?\s*?</\1>')
    topLinkPattern = re.compile(u'(?i)<p.*?>(\u9875\u9996|top)</p>')   # \u9875\u9996 eques 'Top' in Chinese
    if os.path.exists(nodejsFile):
        with io.open(nodejsFile, 'r', encoding='utf-8') as fh:
            content = fh.read()
            htmls = re.findall(r'(?is) ["\']([^"\']*?\.html)["\']', content)
    clearHtmls = []
    for html in htmls:
        if html in clearHtmls:
            continue
        elif any([re.match(ex, html) for ex in EXCLUDELIST]):
            # print u'---------------{}'.format(html)
            continue
        elif re.match(r'(?i)https?://', html):
            continue
        else:
            clearHtmls.append(html)
    tempFolder = path + '_temp'
    if os.path.exists(tempFolder):
        shutil.rmtree(tempFolder)
        print u'[+] removed old temp folder.'
    shutil.copytree(path, tempFolder)
    print u'[+] copied new temp folder.'
    for html in clearHtmls:  # remove all a tag to get better view in docx
        try:
            with io.open(os.path.join(tempFolder, html), 'r+', encoding='utf-8') as fh:
                content = fh.read()
                content = aPattern.sub(lambda m: m.group(1), content)
                content = blankLinePattern.sub('', content)
                content = topLinkPattern.sub('', content)
                fh.seek(0)
                fh.truncate()
                fh.write(content)
        except IOError as e:
            print '[!] generateDocx() - Could not open file: {}'.format(e)

    print u'[+] Going to convert {} html files to docx'.format(len(clearHtmls))
    threadList = []
    docxList = []
    for index, subHtmls in enumerate([clearHtmls[x: x + 88] for x in xrange(0, len(clearHtmls), 88)]):
        thisDocFileName = os.path.join(tempFolder, str(index) + '.docx')
        thisDocFileName = os.path.realpath(thisDocFileName)
        thisThread = threading.Thread(target=htmlToDoc, args=(subHtmls, tempFolder, thisDocFileName))
        threadList.append(thisThread)
        docxList.append(thisDocFileName)
        thisThread.start()
    for thread in threadList:
        thread.join()

    finalDoc = word.Documents.Add()
    docFileName = os.path.realpath(path) + time.strftime('%H_%M_%S', time.localtime(time.time())) + '.docx'
    print u'[*] Going to marge the finall docx file.'
    for docFile in docxList:
        try:
            finalDoc.Application.Selection.Range.InsertFile(docFile)
            # 3=word.WdBreakType.wdSectionBreakContinuous
            finalDoc.Application.Selection.Range.InsertBreak(3)
            # 6=word.WdUnits.wdStory  0=word.WdMovementType.wdMove
            finalDoc.Application.Selection.EndKey(6, 0)
        except Exception as e:
            print u'[!] Faild to add file: {} - {}'.format(docFile, e)
    try:
        finalDoc.ActiveWindow.View.Type = 3  # 3=Word.WdViewType.wdPrintView
        finalDoc.SaveAs(docFileName, FileFormat=12)
        finalDoc.Close()
        print u'[***] Fianl docx file created: {}'.format(docFileName)
    except Exception as e:
        print u'[!] Faild to save final file: {} - {}'.format(finalDoc, e)
    finally:
        word.Quit()
    shutil.rmtree(tempFolder)
    print u'[+] removed this temp folder.'


def htmlToDoc(htmls, path, docFileName='x.docx'):
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    # word.Visible = 1
    word.NormalTemplate.Saved = 1  # tofix : prompted to save normal.dot when quick word.
    sumDoc = word.Documents.Add()
    errordict = {}
    msgLength = 0
    for index, html in enumerate(htmls):
        filePath = os.path.realpath(os.path.join(path, html))
        try:
            doc = word.Documents.Add(filePath)
            docFile = os.path.realpath(os.path.join(path, html + '.docx'))
            try:
                i = 0
                s = len(doc.InlineShapes)
                if s > LIMITPICNO:
                    errorMsg = u'[!] Passed file: {} pictures in {}, limit count is {}. Pass it.'.format(s, LIMITPICNO, html)
                    errordict[html] = errorMsg
                    print u'{}{}'.format(errorMsg, ' ' * (msgLength - len(errorMsg)))
                    doc.Close()
                    continue
                if s > 100:
                    print u'[...] {} pictures to e save, it will take a little long time, Please wait - {}{}'.format(s, html, ' ' * msgLength)
                while(i < s):
                    # 4=word.WdInlineShapeType.wdInlineShapeLinkedPicture
                    if doc.InlineShapes[i].Type == 4:
                        doc.InlineShapes[i].LinkFormat.Update()
                        link = doc.InlineShapes[i].LinkFormat.SourceFullName
                        msg = ' ----[' + html + '] going to save picture: ' + str(i) + "/" + str(s) + ' -->' + link
                        if len(msg) > msgLength:
                            msgLength = len(msg)
                        print u'{}{}\r'.format(msg, ' ' * msgLength),
                        sys.stdout.flush()
                        doc.InlineShapes[i].LinkFormat.SavePictureWithDocument = True
                    i = i + 1
            except Exception as e:
                errordict[html] = e
                print u'[!] Handle picture problem: {}'.format(' ' * msgLength)
                print e
            doc.SaveAs(docFile, FileFormat=12)
            doc.Close()
            msg = u'[+] Convert doc Success with: {} -- {} / {} - {}'.format(docFileName.split('_temp')[-1], index, len(htmls), html)
            if len(msg) > msgLength:
                msgLength = len(msg)
            print u'{}{}\r'.format(msg, ' ' * msgLength),
            sys.stdout.flush()
        except Exception as e:
            print u'[!] convert doc Error: {} - {}{}'.format(html, e, ' ' * msgLength)
            doc.Close()
        sumDoc.Application.Selection.Range.InsertFile(docFile)
        # 3=word.WdBreakType.wdSectionBreakContinuous
        sumDoc.Application.Selection.Range.InsertBreak(3)
        # 6=word.WdUnits.wdStory  0=word.WdMovementType.wdMove
        sumDoc.Application.Selection.EndKey(6, 0)

    i = 0
    while(i < len(sumDoc.Tables)):
        try:
            # 2=Word.WdAutoFitBehavior.wdAutoFitWindow
            sumDoc.Tables[i].AutoFitBehavior(2)
            i = i + 1
        except Exception as e:
            errordict[html] = e
            print u'[!] Error when set autofit of table:{}'.format(' ' * msgLength)
            print e
    try:
        sumDoc.ActiveWindow.View.Type = 3  # 3=Word.WdViewType.wdPrintView
        sumDoc.SaveAs(docFileName, FileFormat=12)
    except Exception as e:
        print u'[!] Error when save docx file: {}{}'.format(docFileName, ' ' * msgLength)
        print e
        errordict[html] = e
    finally:
        sumDoc.Close()
        word.Quit()
    print u'[*] docx created for sub html file list: {}{}'.format(docFileName, ' ' * msgLength)
    return errordict


def main(path):
    addHelpTemplate(path)  # OpManager help cn  #Applications Manager User Guide cn
    generateDocx(path)


if __name__ == '__main__':
    lang, charset = locale.getdefaultlocale()
    if len(sys.argv) > 1:
        path = ' '.join(sys.argv[1:]).decode(charset)
        if os.path.exists(path):
            main(path)
        else:
            print '\n[!] Specified folder path does not exist.\nUsage:\n python helpIQ.py <exported site\'s folder name>'
    else:
        print 'Usage:\n python helpIQ.py <unzipped site\'s folder path>'
