import os
from shutil import rmtree
from lxml import etree
import zipfile
from pathlib import Path
import re
import uuid
from collections import OrderedDict
import json
import fnmatch

symbol_map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
}

re_p = r'<a:p>(.*?)</a:p>'
re_r = r'<a:r>(.*?)</a:r>'
re_t = r'<a:t>(.*?)</a:t>'


def parser_xml_func(unz_file):
    with unz_file.open(mode='r', encoding='utf-8') as f:
        unz_file_xml = f.read()

    # 解析XML文件
    final_xml_str_list = []
    st_1 = 0
    p_element_list = list(re.finditer(re_p, unz_file_xml))
    for index_p, p_element in enumerate(p_element_list):
        st_p, end_p = p_element.span()[0], p_element.span()[1]
        final_xml_str_list.append(unz_file_xml[st_1: st_p])
        r_element_list = list(re.finditer(re_r, p_element.group()))
        if len(r_element_list) < 1:
            st_2 = 0
            t_element_list = list(re.finditer(re_t, p_element.group()))
            for index_t, t_element in enumerate(t_element_list):
                st_t0, end_t0 = p_element.span()[0], p_element.span()[1]
                final_xml_str_list.append(p_element.group()[st_2: st_t0])
                final_xml_str_list.append(p_element.group()[st_t0: end_t0])
                st_2 = end_t0

                if index_t == len(t_element_list) - 1 and len(p_element.group()[end_t0: ]) != 0:
                    final_xml_str_list.append(p_element.group()[end_t0: ])
        else:
            st_3 = 0
            preindex_r = 0
            for index_r, r_element in enumerate(r_element_list):
                st_r0, end_r0 = r_element.span()[0], r_element.span()[1]
                if len(p_element.group()[st_3: st_r0]) != 0:
                    final_xml_str_list.append(p_element.group()[st_3: st_r0])
                st_3 = end_r0
                wt1 = re.search(re_t, r_element.group())
                wt1_ = re.findall(re_t, r_element.group())
                if wt1 is not None:
                    if index_r > 0 and index_r - preindex_r == 1 and '</a:t>' in final_xml_str_list[-1]:
                        if len(wt1_) > 1:
                            tmp_xml_str = ' '.join([item[-1] for item in wt1_])
                            final_xml_str_list[-1] = final_xml_str_list[-1].split('</a:t>')[0] + tmp_xml_str + '</a:t>' + final_xml_str_list[-1].split('</a:t>')[1]
                        else:
                            final_xml_str_list[-1] = final_xml_str_list[-1].split('</a:t>')[0] + wt1.group().split('<a:t>')[1].split('</a:t>')[0] + '</a:t>' + final_xml_str_list[-1].split('</a:t>')[1]
                    else:
                        if len(wt1_) > 1:
                            tmp_xml_str = ' '.join([item[-1] for item in wt1_])
                            tmp_wt1 = r_element.group().split('<a:t>')[0] + '<a:t>' + tmp_xml_str + '</a:t>' + r_element.group().split('</a:t>')[-1]
                            final_xml_str_list.append(tmp_wt1)
                        else:
                            final_xml_str_list.append(r_element.group()) 
                    preindex_r = index_r
                else:
                    final_xml_str_list.append(r_element.group())
                if index_r == len(r_element_list) - 1 and len(p_element.group()[st_3: ]) != 0:
                    final_xml_str_list.append(p_element.group()[st_3: ])
        st_1 = end_p
        if index_p == len(p_element_list) - 1 and len(unz_file_xml[st_1: ]) != 0:
            final_xml_str_list.append(unz_file_xml[st_1: ])

    return final_xml_str_list


# pptx文档解压
def pptx_unzip_and_parser(pptx_path):
    pptx_path = Path(pptx_path) if isinstance(pptx_path, str) else pptx_path
    upzip_path = pptx_path.with_name(pptx_path.stem)
    with zipfile.ZipFile(pptx_path, 'r') as f:
        for file in f.namelist():
            f.extract(file, path=upzip_path)
    unzip_file_dict = dict()
    pattern = "slide*.xml"
    upzip_files = [unz_file for unz_file in upzip_path.glob('ppt/slides/*.*') if fnmatch.fnmatch(str(unz_file.name), pattern)]
    for unz_file in upzip_files:
        final_xml_str_list = parser_xml_func(unz_file)
        if len(final_xml_str_list) > 0:
            unzip_file_dict[str(unz_file)] = final_xml_str_list

    return unzip_file_dict, upzip_path


# 讲文件夹中的所有文件压缩成pptx文档
def pptx_zipped(pptx_path, zipped_path):
    pptx_path = Path(pptx_path) if isinstance(pptx_path, str) else pptx_path
    with zipfile.ZipFile(zipped_path, 'w', zipfile.zlib.DEFLATED) as f:
        for file in pptx_path.glob('**/*.*'):
            f.write(file, file.as_posix().replace(pptx_path.as_posix() + '/', ''))
    zipped_path = pptx_path.parent.joinpath(zipped_path)
    return zipped_path

# 删除生成的解压文件夹
def remove_folder(path):
    path = Path(path) if isinstance(path, str) else path
    if path.exists():
        rmtree(path)
    else:
        raise "系统找不到指定的文件"

def get_pptx_text_per_para(unzip_file_dict, unzip_path, fileuuid):
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
            wt = re.search(re_t, item)
            if wt is not None:
                tmp_wt = wt.group().replace('<a:t>', '').replace('</a:t>', '').strip()
                if len(tmp_wt) < 1:
                    continue
                text_dict[uuid4_] = tmp_wt

    with open(fileuuid, 'w', encoding='utf-8') as fw:
        fw.write(json.dumps(new_unzip_file_dict, ensure_ascii=False))
    return fileuuid, text_dict


def parser_pptx_trans_result(trans_result, fileuuid, pptxname):
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
                file_info[filename][key] = file_info[filename][key].split('<a:t>')[0] + '<a:t>' + trans_result[key] + '</a:t>' + file_info[filename][key].split('</a:t>')[1]

        final_xml_str = ''.join([item for key, item in file_info[filename].items()])

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(final_xml_str)
    
    zipped_path = pptx_zipped(unzip_path, pptxname)
    # remove_folder(unzip_path)

    return zipped_path
      