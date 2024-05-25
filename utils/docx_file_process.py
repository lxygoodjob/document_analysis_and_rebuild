import os
from shutil import rmtree
from lxml import etree
import zipfile
from pathlib import Path
import re
import uuid
from collections import OrderedDict
import json
# from concurrent.futures import ThreadPoolExecutor, as_completed


symbol_map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
}

re_b = r'<w:body>(.*?)</w:body>'
re_wp1 = r'<w:p((?!ict|Pr|Style|:inline|roofErr|gSz|gMar).*?)>(.*?)'
re_wp2 = r'<w:p((?!ict|Pr|Style|:inline|roofErr|gSz|gMar).*?)>(.*?)</w:p>'
re_wr = r'<w:r((?!Pr|Fonts).*?)>(.*?)</w:r>'
re_wt1 = r'<w:t((?!xbxContent|ag w:val=""/|ext|bl|r|d|c|op|a).*?)>'
re_wt2 = r'<w:t((?!xbxContent|ag w:val=""/|ext|bl|r|d|c|op|a).*?)>(.*?)</w:t>'


def parser_xml_func(unz_file):
    with unz_file.open(mode='r', encoding='utf-8') as f:
        unz_file_xml = f.read()

    # 解析XML文件
    final_xml_str_list = []
    body_element = re.search(re_b, unz_file_xml)
    if body_element is not None:
        st_b, end_b = body_element.span()[0], body_element.span()[1]
        final_xml_str_list.append(unz_file_xml[:st_b])
        final_xml_str_list.append('<w:body>')

        root = etree.fromstring(unz_file_xml.encode('utf-8'))

        subchild_list = []
        for child in root:
            for subindex, subchild in enumerate(child):
                xml_str = etree.tostring(subchild, encoding='utf-8').decode()
                subchild_list.append(xml_str)

        for index_p, paragraph in enumerate(subchild_list):
            preindex_r = 0
            wp_list = list(re.finditer(re_wp1, paragraph))
            if len(wp_list) < 1:
                final_xml_str_list.append(paragraph)
            elif len(wp_list) == 1:
                wr_list = list(re.finditer(re_wr, paragraph))
                st_1 = 0
                for index_r, wr in enumerate(wr_list):
                    st_wr, end_wr = wr.span()[0], wr.span()[1]
                    if len(paragraph[st_1: st_wr]) != 0:
                        temp = paragraph[st_1: st_wr]
                        if '<w:proofErr' in temp:
                            proofErr = list(re.finditer(r'<w:proofErr w:type=(.*?)/>', paragraph))
                            for er in proofErr:
                                temp = temp.replace(er.group(), '')
                        if len(temp) > 0:
                            final_xml_str_list.append(temp)
                    st_1 = end_wr
                    wt = re.search(re_wt2, wr.group())
                    wt_ = re.findall(re_wt2, wr.group())

                    if wt is not None:
                        if index_r > 0 and index_r - preindex_r == 1 and '</w:t>' in final_xml_str_list[-1]:
                            if len(wt_) > 1:
                                tmp_xml_str = ' '.join([item[-1] for item in wt_])
                                final_xml_str_list[-1] = final_xml_str_list[-1].split('</w:t>')[0] + tmp_xml_str + '</w:t>' + final_xml_str_list[-1].split('</w:t>')[1]
                            else:
                                split_single = re.search(re_wt1, wr.group()).group()
                                final_xml_str_list[-1] = final_xml_str_list[-1].split('</w:t>')[0] + wt.group().split(split_single)[1].split('</w:t>')[0] + '</w:t>' + final_xml_str_list[-1].split('</w:t>')[1]
                        else:
                            split_single = re.search(re_wt1, wr.group()).group()
                            if len(wt_) > 1:
                                tmp_xml_str = ' '.join([item[-1] for item in wt_])
                                tmp_wr = wr.group().split(split_single)[0] + split_single + tmp_xml_str + '</w:t>' + wr.group().split('</w:t>')[-1]
                                final_xml_str_list.append(tmp_wr)
                            else:
                                final_xml_str_list.append(wr.group()) 
                        preindex_r = index_r
                    else:
                        final_xml_str_list.append(wr.group())
                    if index_r == len(wr_list) - 1 and len(paragraph[st_1: ]) != 0:
                        final_xml_str_list.append(paragraph[st_1: ])
            else:
                wp_list = list(re.finditer(re_wp1, paragraph))
                st_2 = 0
                sub_paragraph_list = []
                for index_p1 in range(1, len(wp_list)):
                    wp1 = wp_list[index_p1]
                    st_wp1, end_wp1 = wp1.span()[0], wp1.span()[1]
                    sub_paragraph_list.append(paragraph[st_2: st_wp1])
                    st_2 = st_wp1
                    if index_p1 == len(wp_list) - 1:
                        sub_paragraph_list.append(paragraph[st_wp1:])

                preindex_p_r = 0
                for sub_paragraph in sub_paragraph_list:
                    if '</w:p>' not in sub_paragraph or '</w:t>' not in sub_paragraph:
                        final_xml_str_list.append(sub_paragraph)
                    else:
                        wp_r_list = list(re.finditer(re_wr, sub_paragraph))
                        st_3 = 0
                        for index_p_r, wp_r in enumerate(wp_r_list):
                            st_wp_r, end_wp_r = wp_r.span()[0], wp_r.span()[1]
                            if len(paragraph[st_3: st_wp_r]) != 0:
                                temp_wpr = sub_paragraph[st_3: st_wp_r]
                                if '<w:proofErr' in temp_wpr:
                                    proofErr = list(re.finditer(r'<w:proofErr w:type=(.*?)/>', sub_paragraph))
                                    for er in proofErr:
                                        temp_wpr = temp_wpr.replace(er.group(), '')
                                if len(temp_wpr) > 0:
                                    final_xml_str_list.append(temp_wpr)
                            st_3 = end_wp_r
                            wp_r_t = re.search(re_wt2, wp_r.group())
                            wp_r_t_ = re.findall(re_wt2, wp_r.group())

                            if wp_r_t is not None:
                                if index_p_r > 0 and index_p_r - preindex_p_r == 1 and '</w:t>' in final_xml_str_list[-1]:
                                    if len(wp_r_t_) > 1:
                                        tmp_wpr_xml_str = ' '.join([item[-1] for item in wp_r_t_])
                                        final_xml_str_list[-1] = final_xml_str_list[-1].split('</w:t>')[0] + tmp_wpr_xml_str + '</w:t>' + final_xml_str_list[-1].split('</w:t>')[1]
                                    else:
                                        split_single = re.search(re_wt1, wp_r_t.group()).group()
                                        final_xml_str_list[-1] = final_xml_str_list[-1].split('</w:t>')[0] + wp_r_t.group().split(split_single)[1].split('</w:t>')[0] + '</w:t>' + final_xml_str_list[-1].split('</w:t>')[1]
                                else:
                                    split_single = re.search(re_wt1, wp_r_t.group()).group()
                                    if len(wp_r_t_) > 1:
                                        tmp_wpr_xml_str = ' '.join([item[-1] for item in wp_r_t_])
                                        tmp_wpr = wp_r_t.group().split(split_single)[0] + split_single + tmp_wpr_xml_str + '</w:t>' + wp_r_t.group().split('</w:t>')[-1]
                                        final_xml_str_list.append(tmp_wpr)
                                    else:
                                        final_xml_str_list.append(wp_r.group()) 
                                preindex_p_r = index_p_r
                            else:
                                final_xml_str_list.append(wp_r.group())
                            if index_p_r == len(wp_r_list) - 1 and len(sub_paragraph[st_3: ]) != 0:
                                final_xml_str_list.append(sub_paragraph[st_3: ])

        final_xml_str_list.append('</w:body>')
        final_xml_str_list.append(unz_file_xml[end_b:])
    
    return final_xml_str_list


# docx文档解压
def docx_unzip_and_parser(docx_path):
    docx_path = Path(docx_path) if isinstance(docx_path, str) else docx_path
    upzip_path = docx_path.with_name(docx_path.stem)
    with zipfile.ZipFile(docx_path, 'r') as f:
        for file in f.namelist():
            f.extract(file, path=upzip_path)
    unzip_file_dict = dict()
    upzip_files = [unz_file for unz_file in upzip_path.glob('**/*.*') if str(unz_file).endswith('document.xml')]

    for unz_file in upzip_files:
        final_xml_str_list = parser_xml_func(unz_file)
        if len(final_xml_str_list) > 0:
            unzip_file_dict[str(unz_file)] = final_xml_str_list

    return unzip_file_dict, upzip_path


# 讲文件夹中的所有文件压缩成docx文档
def docx_zipped(docx_path, zipped_path):
    docx_path = Path(docx_path) if isinstance(docx_path, str) else docx_path
    with zipfile.ZipFile(zipped_path, 'w', zipfile.zlib.DEFLATED) as f:
        for file in docx_path.glob('**/*.*'):
            f.write(file, file.as_posix().replace(docx_path.as_posix() + '/', ''))
    zipped_path = docx_path.parent.joinpath(zipped_path)
    return zipped_path

# 删除生成的解压文件夹
def remove_folder(path):
    path = Path(path) if isinstance(path, str) else path
    if path.exists():
        rmtree(path)
    else:
        raise "系统找不到指定的文件"

def get_docx_text_per_para(unzip_file_dict, unzip_path, fileuuid):
    new_unzip_file_dict = dict()
    new_unzip_file_dict['unzip_path'] = str(unzip_path)
    new_unzip_file_dict['file_info'] = dict()
    text_dict = dict()
    for file_name in unzip_file_dict.keys():
        final_xml_str_list = unzip_file_dict[file_name]
        new_unzip_file_dict['file_info'][file_name] = OrderedDict()
        for index_f, item in enumerate(final_xml_str_list):
            uuid4_ = str(uuid.uuid4())
            new_unzip_file_dict['file_info'][file_name][uuid4_] = item
            # wt = re.search(r'<w:t(.*?)>(.*?)</w:t>', item)
            wt = re.search(re_wt2, item)
            if wt is not None:
                # split_single = re.search(r'<w:t(.*?)>', item).group()
                split_single = re.search(re_wt1, item).group()
                # print(11111, uuid4_, split_single)
                # if re.search(u'[\u4e00-\u9fa5]', wt.group()):
                tmp_wt = wt.group().replace(split_single, '').replace('</w:t>', '').strip()
                if len(tmp_wt) < 1:
                    continue
                text_dict[uuid4_] = tmp_wt

    with open(fileuuid, 'w', encoding='utf-8') as fw:
        fw.write(json.dumps(new_unzip_file_dict, ensure_ascii=False))
    return fileuuid, text_dict


def parser_docx_trans_result(trans_result, fileuuid, docxname):
    with open(fileuuid, 'r', encoding='utf-8') as fr:
        new_unzip_file_dict = json.loads(fr.read())
    unzip_path = new_unzip_file_dict['unzip_path']
    file_info = new_unzip_file_dict['file_info']
    for filename in file_info.keys():
        for key in trans_result.keys():
            if key in file_info[filename].keys():
                for skey in symbol_map.keys():
                    if symbol_map[skey] in trans_result[key]:
                        continue
                    else:
                        trans_result[key] =  trans_result[key].replace(skey, symbol_map[skey])
                split_single = re.search(re_wt1, file_info[filename][key]).group()
                file_info[filename][key] = file_info[filename][key].split(split_single)[0] + split_single + trans_result[key] + '</w:t>' + file_info[filename][key].split('</w:t>')[1]

        final_xml_str = ''.join([item for key, item in file_info[filename].items()])

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(final_xml_str)
    
    zipped_path = docx_zipped(unzip_path, docxname)
    # remove_folder(unzip_path)

    return zipped_path
      
if __name__ == '__main__':
    unzip_file_dict, unzip_path = docx_unzip_and_parser('/opt/xiaoyu/DT/files/GEN AI POC-EN.docx')
    print(unzip_file_dict)