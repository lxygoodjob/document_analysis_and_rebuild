import os
from langchain_community.document_loaders import DirectoryLoader
from typing import Any, List, Optional, Type, Union
from .chinese_recursive_text_splitter import ChineseRecursiveTextSplitter
from pathlib import Path
import uuid

def _is_visible(p: Path) -> bool:
    parts = p.parts
    for _p in parts:
        if _p.startswith("."):
            return False
    return True

def split_func(file, chunk_size=300, chunk_overlap=0, keep_separator=True, is_separator_regex=True):
    filename = os.path.basename(file)
    path = os.path.dirname(file)
    loader = DirectoryLoader(path)
    p = Path(path)
    file_path = p.joinpath(filename)
    documents = list(loader._lazy_load_file(file_path, p, None))
    text_splitter = ChineseRecursiveTextSplitter(keep_separator=keep_separator, is_separator_regex=is_separator_regex, chunk_size=chunk_size, chunk_overlap=chunk_overlap)
    split_docs = text_splitter.split_documents(documents)
    contents = dict()
    for i in split_docs:
        uuid4_ = str(uuid.uuid4())
        contents[uuid4_] = i.page_content
    return contents