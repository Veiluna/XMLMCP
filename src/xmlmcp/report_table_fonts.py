"""Read the fonts actually selected for text in placeable WMF reference tables.

MS-WMF: records are measured in WORDs; GDI objects occupy the lowest free slot.
This bounded reader does not render metafiles or execute their drawing commands.
https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/6040492f-7b58-49bd-bfef-ef1126bdffe3
"""
import struct
from collections import Counter

from .errors import require
from .namespaces import A, XDR


def wmf_text_fonts(data):
    require(len(data) >= 40 and data[:4] == b'\xd7\xcd\xc6\x9a', 'TEMPLATE_FONT', '参考表格不是支持的 placeable WMF')
    left, top, right, bottom, units = struct.unpack_from('<4hH', data, 6)
    require(right > left and bottom > top and units > 0, 'TEMPLATE_FONT', 'WMF 参考表格边界无效')
    position, objects, selected, texts, eof = 40, {}, None, [], False
    while position + 6 <= len(data):
        words, function = struct.unpack_from('<IH', data, position)
        require(words >= 3 and position + words * 2 <= len(data), 'TEMPLATE_FONT', 'WMF 记录越界')
        payload = data[position + 6:position + words * 2]
        if function in (0x2fa, 0x2fb, 0x2fc, 0x1f9, 0xf7, 0x6ff, 0x142):
            index = next(i for i in range(65536) if i not in objects)
            objects[index] = {'kind': function}
            if function == 0x2fb:
                require(len(payload) >= 18, 'TEMPLATE_FONT', 'WMF 字体记录无效')
                objects[index].update(height=struct.unpack_from('<h', payload)[0],
                                      weight=struct.unpack_from('<h', payload, 8)[0],
                                      italic=bool(payload[10]), underline=bool(payload[11]),
                                      face=payload[18:].split(b'\0')[0].decode('gb18030'))
        elif function == 0x1f0:
            require(len(payload) >= 2, 'TEMPLATE_FONT', 'WMF 删除对象记录无效')
            objects.pop(struct.unpack_from('<H', payload)[0], None)
        elif function == 0x12d:
            require(len(payload) >= 2, 'TEMPLATE_FONT', 'WMF 选择对象记录无效')
            obj = objects.get(struct.unpack_from('<H', payload)[0])
            if obj and obj['kind'] == 0x2fb:
                selected = obj.copy()
        elif function == 0xa32:
            require(len(payload) >= 8, 'TEMPLATE_FONT', 'WMF 文字记录无效')
            y, x, count, options = struct.unpack_from('<hhHH', payload)
            start = 16 if options & 6 else 8
            require(start + count <= len(payload), 'TEMPLATE_FONT', 'WMF 文字长度越界')
            text = payload[start:start + count].decode('gb18030')
            if text.strip():
                require(selected is not None, 'TEMPLATE_FONT', 'WMF 文字缺少选定字体')
                texts.append({'x': x, 'y': y, 'text': text, 'font': selected.copy()})
        position += words * 2
        if function == 0:
            eof = True
            break
    require(eof, 'TEMPLATE_FONT', 'WMF 缺少完整结束记录')
    return bottom - top, texts


def table_fonts(reference_book, shape_id, single):
    matches = [i for i in reference_book.images() if i['shape_id'] == str(shape_id)]
    if not matches:
        # Synthetic fixtures and older image-only mappings have no extractable fonts.
        return {r: {'name': '宋体', 'size_pt': 9, 'bold': not single and r != 'body',
                    'italic': False, 'underline': False, 'basis': 'explicit fallback'}
                for r in ('title', 'header', 'body')}
    require(len(matches) == 1, 'TEMPLATE_FONT', '参考表格图片不唯一')
    image = matches[0]
    data = reference_book.package.parts[image['media_part']]
    height, texts = wmf_text_fonts(data)
    transform = image['_pic'].child('spPr', XDR).child('xfrm', A)
    require(transform is not None and transform.child('ext', A) is not None, 'TEMPLATE_FONT', '参考表格没有明确尺寸')
    scale = int(transform.child('ext', A).attrs['cy']) / 12700 / height
    ys = sorted({t['y'] for t in texts})
    baselines = []
    for y in ys:
        # Latin/symbol fallback runs can sit one or two logical pixels above Chinese.
        if baselines and y - baselines[-1][-1] <= 2:
            baselines[-1].append(y)
        else:
            baselines.append([y])
    require(len(baselines) >= 3, 'TEMPLATE_FONT', '参考表格缺少标题、表头和数据字体')
    result = {}
    for role, baseline in zip(('title', 'header', 'body'), baselines[:3]):
        role_texts = [t for t in texts if t['y'] in baseline]
        chinese_texts = [t for t in role_texts if any('\u3400' <= c <= '\u9fff' for c in t['text'])]
        fonts = [t['font'] for t in chinese_texts or role_texts]
        signatures = [(f['face'], f['height'], f['weight'], f['italic'], f['underline']) for f in fonts]
        signature = Counter(signatures).most_common(1)[0][0]
        name, logical_height, weight, italic, underline = signature
        require(logical_height < 0, 'TEMPLATE_FONT', '仅支持明确字符高度的参考字体')
        result[role] = {'name': name, 'size_pt': round(abs(logical_height) * scale, 2),
                        'bold': weight > 400, 'italic': italic, 'underline': underline,
                        'wmf_weight': weight, 'wmf_height': logical_height, 'basis': 'WMF selected text font and original picture height'}
    cov_fonts = [t['font'] for t in texts if 'COV' in t['text']]
    if single and cov_fonts:
        f = cov_fonts[0]
        result['cov_label'] = {'name':f['face'], 'size_pt':round(abs(f['height'])*scale,2), 'bold':f['weight']>400,
                               'italic':f['italic'], 'underline':f['underline'], 'wmf_weight':f['weight'], 'wmf_height':f['height'],
                               'basis':'WMF COV label font and original picture height'}
    return result
