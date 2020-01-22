import re
import zipfile


def open_file(filename):
    text = open(filename, encoding='utf-8').read()
    return text


def elan(filename):
    elan = open_file(filename)
    elan = elan.replace('&', '')
    elan = elan.replace('<', '&lt;')
    elan = elan.replace('>', '&gt;')
    elan = elan.splitlines()
    transc = []
    transl = []
    gloss = []
    comment = []
    for line in elan:
        tokens = line.split('\t')
        layer = tokens[0]
        time_start = tokens[2]
        time_finish = tokens[4]
        text = tokens[8]
        if layer == 'transcription':
            transc.append([text, time_start, time_finish])
        elif layer == 'translation':
            transl.append([text, time_start, time_finish])
        elif layer == 'gloss':
            gloss.append([text, time_start, time_finish])
        elif layer == 'comment':
            comment.append([text, time_start, time_finish])
    return transc, transl, gloss, comment


def small_caps(text):
    new_text = re.sub('<.+?>', ' ', text)
    pattern = '[a-z-=]+'
    latins = re.findall(pattern, new_text)
    for latin in latins:
        text = text.replace(latin, '</w:t></w:r><w:r w:rsidRPr="00F6391B"><w:rPr><w:smallCaps/><w:lang w:val="en-US"/>'
                                   '</w:rPr><w:t>' + latin + '</w:t></w:r><w:r><w:t>')
    return text


def write_to_word(transc, transl, gloss, comment):
    print(len(transl), len(transc), len(gloss), len(comment))
    length = len(transc)
    informant = input('введите код информанта ')
    data = input('введите дату ')
    expe = input('введите свой код ')
    name = f'eve_{informant}_{data}_{expe}.docx'
    to_write = []
    for i in range(length):
        part = open_file('tag.txt')
        part = part.replace('informant', informant)
        part = part.replace('data', data)
        part = part.replace('expe', expe)
        part = part.replace('number', str(i + 1))
        if transc[i][1] == transl[i][1] or transc[i][2] == transl[i][2]:
            if gloss[i][1] == transc[i][1] or gloss[i][2] == transc[i][2]:
                part = part.replace('glossing',
                                    '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                                        gloss[i][0].split()))
                part = small_caps(part)
            else:
                gloss.insert(i, ['', '0', '0'])
                part = part.replace('glossing',
                                    '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                                        gloss[i][0].split()))
                part = small_caps(part)
            part = part.replace('TEXT', '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                transc[i][0].split()))
            part = part.replace('translation', '</w:t></w:r><w:r><w:t>' + transl[i][0])
        else:
            transl.insert(i, ['', '0', '0'])
            if gloss[i][1] == transc[i][1] or gloss[i][2] == transc[i][2]:
                part = part.replace('glossing',
                                    '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                                        gloss[i][0].split()))
                part = small_caps(part)
            else:
                gloss.insert(i, ['', '0', '0'])
                part = part.replace('glossing',
                                    '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                                        gloss[i][0].split()))
                part = small_caps(part)
            part = part.replace('TEXT', '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                transc[i][0].split()))
            part = part.replace('translation', '</w:t></w:r><w:r><w:t>' + transl[i][0])
        try:
            if comment[i][1] == transc[i][1] or comment[i][2] == transc[i][2]:
                part = part.replace('optional', f'</w:t></w:r><w:r><w:t>{transc[i][1]}—{transc[i][2]} {comment[i][0]}')
            else:
                part = part.replace('optional', f'{transc[i][1]}—{transc[i][2]}')
                comment.insert(i, ['', '0', '0'])
        except Exception:
            part = part.replace('optional', f'{transc[i][1]}—{transc[i][2]}')
        to_write.append(part)
    docx = open_file('document1.xml')
    docx = docx.replace('PASTE_HERE', ''.join(to_write))
    with open('document.xml', 'w', encoding='utf-8') as f:
        f.write(docx)
    return name


def new_word(name):
    y = zipfile.ZipFile(name, 'w')
    z = zipfile.ZipFile('sample.docx', 'r')
    for file in z.namelist():
        if file == 'word/document.xml':
            y.write('document.xml', arcname='word/document.xml')
        else:
            data = z.read(file)
            y.writestr(file, data)


def main():
    file = input('введите название илановского файла или назовите его 1.txt и нажмите Enter')
    if file == '':
        file = '1.txt'
    el = elan(file)
    transc = el[0]
    transl = el[1]
    gloss = el[2]
    comment = el[3]
    name = write_to_word(transc, transl, gloss, comment)
    new_word(name)


if __name__ == '__main__':
    main()
