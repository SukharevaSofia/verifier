"""
"""

import struct
from . import xolefile as OLE
from enum import Enum


class StructInt(Enum):
    U8 = '<B'
    U16 = '<H'
    U32 = '<I'


class DocExceptions(Enum):
    NOFILE = "Ошибка чтения файла"
    CORRUPT = "Повреждённый файл"
    NOSUPPORT = "Файл не поддерживается"


class DocFileType(Enum):
    DOC1997 = 1997
    DOC2000 = 2000
    DOC2002 = 2002
    DOC2003 = 2003
    DOC2007 = 2007


def bits(in_bytes):
    for b in in_bytes:
        for offset in range(8):
            yield (b >> offset) & 1


def parse_sttb(data):
    fExtend = data[0:2]
    fExtend = fExtend == bytes.fromhex('ffff')

    result = {"extended": fExtend, "data": [], "extradata": []}

    cchData_format = StructInt.U16.value if fExtend else StructInt.U8.value
    cchData_size = 2 if fExtend else 1

    cData_start = 2 if fExtend else 0
    cData_end = cData_start + 2
    cData = struct.unpack(StructInt.U16.value, data[cData_start:cData_end])[0]

    cbExtra_start = cData_end
    cbExtra_end = cbExtra_start + 2
    cbExtra = struct.unpack(StructInt.U16.value,
                            data[cbExtra_start:cbExtra_end])[0]

    last_position = cbExtra_end
    for i in range(cData):
        cchData_start = last_position
        cchData_end = cchData_start + cchData_size
        cchData = struct.unpack(
            cchData_format, data[cchData_start:cchData_end])[0]

        char_size = 2 if fExtend else 1
        Data_start = cchData_end
        Data_end = Data_start + cchData * char_size
        result["data"].append(data[Data_start:Data_end])

        ExtraData_start = Data_end
        ExtraData_end = ExtraData_start + cbExtra
        result["extradata"].append(data[ExtraData_start: ExtraData_end])
        last_position = ExtraData_end

    return result


def file_type_lookup(cbRgFcLcb):
    # comments specify nFib
    match cbRgFcLcb:
        case 0x005D:  # 0x00C1
            return DocFileType.DOC1997
        case 0x006C:  # 0x00D9
            return DocFileType.DOC2000
        case 0x0088:  # 0x0101
            return DocFileType.DOC2002
        case 0x00A4:  # 0x010C
            return DocFileType.DOC2003
        case 0x00B7:  # 0x0112
            return DocFileType.DOC2007
        case _:
            raise Exception(DocExceptions.CORRUPT, cbRgFcLcb)


def parse_FIB(stream):
    fib_base = stream[0:32]
    wIdent = struct.unpack(StructInt.U16.value, fib_base[0:2])[0]
    if (wIdent != 0xA5EC):
        raise Exception(DocExceptions.CORRUPT)

    flags = fib_base[10:12]

    pnNext = struct.unpack(StructInt.U16.value, fib_base[8:10])[0]

    fComplex = list(bits(flags))[2]
    if (fComplex):
        raise Exception(DocExceptions.NOSUPPORT +
                        "Нет поддержки инкрементальных файлов")

    fHasPic = list(bits(flags))[3] == 1

    fEncrypted = list(bits(flags))[8]
    if (fEncrypted):
        raise Exception(DocExceptions.NOSUPPORT +
                        "Нет поддержки зашифрованных файлов")

    fWhichTblStm = list(bits(flags))[9]
    return {"table": "1Table" if fWhichTblStm else "0Table",
            "pnNext": pnNext,
            "fHasPic": fHasPic}


def spra_size(spra, sprm, nextbyte, next2byte):
    match spra:
        case 0:
            return 1
        case 1:
            return 1
        case 2:
            return 2
        case 3:
            return 4
        case 4:
            return 2
        case 5:
            return 2
        case 6:
            if sprm == 0xD608:
                return struct.unpack(StructInt.U16.value, next2byte)[0] + 1
            if sprm == 0xC615:
                cb = struct.unpack(StructInt.U16.value, next2byte)[0]
                if cb == 255:
                    raise Exception(DocExceptions.NOSUPPORT)
                return cb
            return struct.unpack(StructInt.U8.value, nextbyte)[0]
        case 7:
            return 3
        case _:
            raise Exception(DocExceptions.CORRUPT)


class ResultField(Enum):
    contents = "Состав содержания"
    nav = "Навигация"
    headers = "Заголовки"
    numeration = "Нумерация"

    sections = "Разделы и подразделы"
    abbrevs = "Список сокращений"
    dictionary = "Список терминов"
    introoutro = "Введение и заключение"
    sources = "Список источников"
    addition = "Приложение"

    img = "Иллюстрации"
    tables = "Таблицы"

    page = "Формат листа"
    fontsz = "Размеры полей"
    intervals = "Интервалы текста"
    font = "Шрифт"


def process_ole(ole):
    result = {
        ResultField.contents.value: None,
        ResultField.nav.value: None,
        ResultField.headers.value: None,
        ResultField.numeration.value: True,

        ResultField.sections.value: None,
        ResultField.abbrevs.value: None,
        ResultField.dictionary.value: None,
        ResultField.introoutro.value: None,
        ResultField.sources.value: None,
        ResultField.addition.value: None,

        ResultField.img.value: None,
        ResultField.tables.value: None,

        ResultField.page.value: True,
        ResultField.fontsz.value: None,
        ResultField.intervals.value: None,
        ResultField.font.value: None,
    }
    if not ole.exists("WordDocument"):
        raise Exception(DocExceptions.CORRUPT)

    docdata = None
    with ole.openstream("WordDocument") as doc:
        docdata = doc.read()
    if docdata is None:
        raise Exception(DocExceptions.CORRUPT)

    fib = parse_FIB(docdata)
    result[ResultField.img.value] = fib["fHasPic"]
    if not ole.exists(fib["table"]):
        raise Exception(DocExceptions.CORRUPT)

    tabledata = None
    with ole.openstream(fib["table"]) as tab:
        tabledata = tab.read()
    if tabledata is None:
        raise Exception(DocExceptions.CORRUPT)

    csw_start = 32
    csw_end = csw_start + 2
    csw = struct.unpack(StructInt.U16.value,
                        docdata[csw_start: csw_end])[0]
    if csw != 0x000E:
        raise Exception(DocExceptions.CORRUPT, csw)

    fibRgW_start = csw_end
    fibRgW_end = 34 + (csw * 2)
    fibRgW = docdata[fibRgW_start:fibRgW_end]

    cslw_start = fibRgW_end
    cslw_end = cslw_start + 2
    cslw = struct.unpack(StructInt.U16.value,
                         docdata[cslw_start:cslw_end])[0]
    if cslw != 0x0016:
        raise Exception(DocExceptions.CORRUPT, cslw)

    fibRgLw_start = cslw_end
    fibRgLw_end = fibRgLw_start+(cslw * 4)
    fibRgLw = docdata[fibRgLw_start:fibRgLw_end]

    cbRgFcLcb_start = fibRgLw_end
    cbRgFcLcb_end = cbRgFcLcb_start + 2
    cbRgFcLcb = struct.unpack(
        StructInt.U16.value, docdata[cbRgFcLcb_start:cbRgFcLcb_end])[0]

    fibRgFcLcbBlob_start = cbRgFcLcb_end
    fibRgFcLcbBlob_end = fibRgFcLcbBlob_start + cbRgFcLcb * 8
    fibRgFcLcbBlob = docdata[fibRgFcLcbBlob_start:fibRgFcLcbBlob_end]

    doctype = file_type_lookup(cbRgFcLcb)
    print("Document type: ", doctype)

    fcStshf = struct.unpack(StructInt.U32.value, fibRgFcLcbBlob[8:12])[0]
    lcbStshf = struct.unpack(StructInt.U32.value, fibRgFcLcbBlob[12:16])[0]

    fcPlcfSed_start = 12 * 4
    fcPlcfSed_end = fcPlcfSed_start + 4
    fcPlcfSed = struct.unpack(
        StructInt.U32.value,
        fibRgFcLcbBlob[fcPlcfSed_start:fcPlcfSed_end]
    )[0]

    lcbPlcfSed_start = 13 * 4
    lcbPlcfSed_end = lcbPlcfSed_start + 4
    lcbPlcfSed = struct.unpack(
        StructInt.U32.value,
        fibRgFcLcbBlob[lcbPlcfSed_start:lcbPlcfSed_end]
    )[0]

    if lcbPlcfSed != 0:
        PclfSed = tabledata[fcPlcfSed:fcPlcfSed+lcbPlcfSed]
        if (lcbPlcfSed-4) % 16 != 0:
            raise Exception(DocExceptions.CORRUPT, lcbPlcfSed)
        section_count = int((lcbPlcfSed-4) / 16)
        Seds = PclfSed[section_count * 4 + 4:]
        for sect in range(section_count):
            Sed = Seds[sect * 12:sect*12+12]
            if len(Sed) != 12:
                raise Exception(DocExceptions.CORRUPT, len(Sed))
            fcSepx = struct.unpack(StructInt.U32.value, Sed[2:6])[0]
            Sepx_cb = struct.unpack(
                StructInt.U16.value, docdata[fcSepx:fcSepx+2])[0]
            grpprl = docdata[fcSepx+2:fcSepx+2+Sepx_cb]
            prl_i = 0
            while prl_i < len(grpprl)-2:
                sprm = struct.unpack(StructInt.U16.value,
                                     grpprl[prl_i:prl_i+2])[0]
                spra = sprm >> 13
                operand_sz = spra_size(
                    spra, sprm,
                    grpprl[prl_i+2:prl_i+3],
                    grpprl[prl_i+2:prl_i+4])
                operand = grpprl[prl_i+2:prl_i+2+operand_sz]
                if sprm == 0xB01F:  # sprmSXaPage
                    XA = struct.unpack(StructInt.U16.value, operand)[0]
                    if XA != 11906:
                        result[ResultField.page.value] = False
                if sprm == 0xB020:  # sprmSYaPage
                    YA = struct.unpack(StructInt.U16.value, operand)[0]
                    if YA != 16838:
                        result[ResultField.page.value] = False
                if sprm == 0x300E:  # sprmSNfcPgn
                    fmt = struct.unpack(StructInt.U8.value, operand)[0]
                    if fmt != 0:
                        result[ResultField.numeration.value] = False
                prl_i = prl_i+2+operand_sz

    fcSttbfFfn_start = 30 * 4
    fcSttbfFfn_end = fcSttbfFfn_start + 4
    fcSttbfFfn = struct.unpack(
        StructInt.U32.value,
        fibRgFcLcbBlob[fcSttbfFfn_start:fcSttbfFfn_end])[0]

    lcbSttbfFfn_start = fcSttbfFfn_end
    lcbSttbfFfn_end = lcbSttbfFfn_start + 4
    lcbSttbfFfn = struct.unpack(
        StructInt.U32.value,
        fibRgFcLcbBlob[lcbSttbfFfn_start:lcbSttbfFfn_end])[0]

    SttbfFfn = tabledata[fcSttbfFfn:fcSttbfFfn + lcbSttbfFfn]
    document_fonts_list = parse_sttb(SttbfFfn)
    document_fonts_list = [font[39:].decode('utf-16le').split("\x00")[0]
                           for font in document_fonts_list["data"]]
    print("Document fonts:", document_fonts_list)

    basetext = docdata.decode('utf-16le', errors='replace')
    basetext_upper = basetext.upper()

    result[ResultField.abbrevs.value] = \
        "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ" in basetext_upper
    result[ResultField.introoutro.value] = \
        "ВВЕДЕНИЕ" in basetext_upper and "ЗАКЛЮЧЕНИЕ" in basetext_upper
    result[ResultField.sources.value] = \
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in basetext_upper
    result[ResultField.addition.value] = \
        "ПРИЛОЖЕНИЕ" in basetext_upper
    result[ResultField.dictionary.value] = \
        "СПИСОК ТЕРМИНОВ И ОПРЕДЕЛЕНИЙ" in basetext_upper

    STSH = tabledata[fcStshf:fcStshf+lcbStshf]
    cbStshi = struct.unpack(StructInt.U16.value, STSH[0:2])[0]
    STSHI = STSH[2:2 + cbStshi]
    stshif = STSHI[0:18]
    stshif_cbSTDBaseInFile = struct.unpack(
        StructInt.U16.value, stshif[2:4])[0]
    if stshif_cbSTDBaseInFile != 10 and stshif_cbSTDBaseInFile != 18:
        raise Exception(DocExceptions.CORRUPT)

    ftcAsci = struct.unpack(StructInt.U16.value, stshif[12:14])[0]
    result[ResultField.font.value] = \
        document_fonts_list[ftcAsci] == "Times New Roman"

    rglpstd = STSH[2 + cbStshi:]

    style_pos = 0
    while style_pos < len(rglpstd):
        if style_pos & 1 == 1:
            style_pos = style_pos + 1
        cbStd = struct.unpack(StructInt.U16.value,
                              rglpstd[style_pos:style_pos+2])[0]
        if cbStd == 0:
            style_pos = style_pos+2
            continue
        std = rglpstd[style_pos+2:style_pos+2+cbStd]
        style_pos = style_pos+2+cbStd

        stdfBase = std[0:10]
        StdfPost2000OrNone = None
        if stshif_cbSTDBaseInFile == 18:
            StdfPost2000OrNone = std[10:18]

        stiABCD = struct.unpack(StructInt.U16.value, stdfBase[0:2])[0]
        sti = stiABCD >> 4

        stk_istdBase = struct.unpack(StructInt.U16.value, stdfBase[2:4])[0]
        stk = stk_istdBase & 0xf
        istdBase = (stk_istdBase & 0xfff0) >> 4

        cupx_istdNext = struct.unpack(
            StructInt.U16.value, stdfBase[4:6])[0]
        cupx = cupx_istdNext & 0xf
        istdNext = (cupx_istdNext & 0xfff0) >> 4

        bchUpe = struct.unpack(StructInt.U16.value, stdfBase[6:8])[0]
        if bchUpe != cbStd:
            raise Exception(DocExceptions.CORRUPT, bchUpe, cbStd)
        grfstd = struct.unpack(StructInt.U16.value, stdfBase[8:10])[0]

        xstzName_cch = struct.unpack(
            StructInt.U16.value,
            std[stshif_cbSTDBaseInFile:stshif_cbSTDBaseInFile+2])[0]
        xstzName_rgtchar_begin = stshif_cbSTDBaseInFile + 2
        xstzName_rgtchar_end = stshif_cbSTDBaseInFile+2+xstzName_cch*2
        xstzName_rgtchar = std[
            xstzName_rgtchar_begin:
            xstzName_rgtchar_end].decode(
            'utf-16le')
        xstzName_chTerm = std[xstzName_rgtchar_end:xstzName_rgtchar_end+2]
        if not xstzName_chTerm == bytes.fromhex('0000'):
            raise Exception(DocExceptions.CORRUPT)

        if stk == 1:  # Paragraph style
            try:
                StkParaGRLPUPX_begin = xstzName_rgtchar_end+2
                LPUpxPapx_cbUpx = struct.unpack(
                    StructInt.U16.value,
                    std[
                        StkParaGRLPUPX_begin:
                        StkParaGRLPUPX_begin+2])[0]
                UpxPapx_begin = StkParaGRLPUPX_begin+2
                UpxPapx_end = UpxPapx_begin + LPUpxPapx_cbUpx
                UpxPapx_len = UpxPapx_end - UpxPapx_begin
                UpxPapx = std[UpxPapx_begin:UpxPapx_end]
                istd = struct.unpack(StructInt.U16.value, UpxPapx[0:2])[0]
                UpxPapx_i = 0
                while UpxPapx_i < UpxPapx_len-2:
                    sprm = struct.unpack(StructInt.U16.value,
                                         UpxPapx[UpxPapx_i:UpxPapx_i+2])[0]
                    spra = sprm >> 13
                    sgc = (0x1c00 & sprm) >> 10
                    fspec = (0x0200 & sprm) >> 9
                    ispmd = (0x01ff & sprm)

                    operand_sz = spra_size(
                        spra,
                        sprm,
                        UpxPapx[UpxPapx_i+2:UpxPapx_i+2+1],
                        UpxPapx[UpxPapx_i+2:UpxPapx_i+2+2])
                    operand = UpxPapx[UpxPapx_i+2:UpxPapx_i+2+operand_sz]

                    if sprm == 0x6412:  # sprmPDyaLine
                        dyaLine = struct.unpack(
                            StructInt.U16.value, operand[0:2])[0]
                        fMultLinespace = struct.unpack(
                            StructInt.U16.value, operand[2:4])[0]
                        key = ResultField.intervals.value
                        if (dyaLine in [240, 360]) and fMultLinespace == 1:
                            if result[key] is None:
                                result[key] = True
                        else:
                            result[key] = False
                    if ispmd == 0x5E:  # sprmPDxaLeft
                        indent = struct.unpack(StructInt.U16.value, operand)[0]
                        if not (140 <= indent and indent <= 144):
                            result[ResultField.intervals.value] = False

                    UpxPapx_i = UpxPapx_i+2+operand_sz
            except Exception:
                pass
        elif stk == 2:  # Char style
            try:
                StkCharGRLPUPX_begin = xstzName_rgtchar_end+2
                LPUpxChpx_cbUpx = struct.unpack(
                    StructInt.U16.value,
                    std[
                        StkCharGRLPUPX_begin:
                        StkCharGRLPUPX_begin+2])[0]
                UpxChpx_begin = StkParaGRLPUPX_begin+2
                UpxChpx_end = UpxChpx_begin + LPUpxChpx_cbUpx
                UpxChpx_len = UpxChpx_end - UpxChpx_begin
                UpxChpx = std[UpxChpx_begin:UpxChpx_end]
                UpxChpx_i = 0
                while UpxChpx_i < UpxChpx_len-2:
                    sprm = struct.unpack(StructInt.U16.value,
                                         UpxChpx[UpxChpx_i:UpxChpx_i+2])[0]
                    spra = sprm >> 13
                    sgc = (0x1c00 & sprm) >> 10
                    fspec = (0x0200 & sprm) >> 9
                    ispmd = (0x01ff & sprm)

                    operand_sz = spra_size(
                        spra,
                        sprm,
                        UpxChpx[UpxChpx_i+2:UpxChpx_i+2+1],
                        UpxChpx[UpxChpx_i+2:UpxChpx_i+2+2])
                    operand = UpxChpx[UpxChpx_i+2:UpxChpx_i+2+operand_sz]

                    if ispmd == 0x43:  # sprmCHps
                        fontsz = struct.unpack(StructInt.U16.value, operand)[0]
                        result[ResultField.fontsz.value] = fontsz == 28

                    UpxChpx_i = UpxChpx_i+2+operand_sz
            except Exception:
                pass
        elif stk == 3:  # Table style
            try:
                StkTableGRLPUPX_begin = xstzName_rgtchar_end+2
                LPUpxTapx_cbUpx = struct.unpack(
                    StructInt.U16.value,
                    std[
                        StkTableGRLPUPX_begin:
                        StkTableGRLPUPX_begin + 2])[0]
                UpxTapx_begin = StkParaGRLPUPX_begin+2
                UpxTapx_end = UpxTapx_begin + LPUpxTapx_cbUpx
                UpxTapx_len = UpxTapx_end - UpxTapx_begin
                UpxTapx = std[UpxTapx_begin:UpxTapx_end]
                UpxTapx_i = 0
                while UpxTapx_i < UpxTapx_len-2:
                    sprm = struct.unpack(StructInt.U16.value,
                                         UpxTapx[UpxTapx_i:UpxTapx_i+2])[0]
                    spra = sprm >> 13
                    sgc = (0x1c00 & sprm) >> 10
                    fspec = (0x0200 & sprm) >> 9
                    ispmd = (0x01ff & sprm)

                    operand_sz = spra_size(
                        spra,
                        sprm,
                        UpxTapx[UpxTapx_i+2:UpxTapx_i+2+1],
                        UpxTapx[UpxTapx_i+2:UpxTapx_i+2+2])
                    operand = UpxTapx[UpxTapx_i+2:UpxTapx_i+2+operand_sz]

                    if sprm == 0x9601:  # sprmTDxaLeft
                        indent = struct.unpack(StructInt.U16.value, operand)[0]
                        result[ResultField.tables.value] = indent == 0

                    UpxTapx_i = UpxTapx_i+2+operand_sz
            except Exception:
                pass
        elif stk == 4:
            pass
        else:
            raise Exception(DocExceptions.CORRUPT, stk)

    print("Opened OLE:" + ole.get_rootentry_name())
    print(ole.listdir())
    if result[ResultField.tables.value] is not False:
        result[ResultField.tables.value] = True

    result[ResultField.sections.value] = \
        result[ResultField.abbrevs.value] and \
        result[ResultField.dictionary.value] and \
        result[ResultField.introoutro.value] and \
        result[ResultField.sources.value] and \
        result[ResultField.addition.value]

    result[ResultField.contents.value] = \
        result[ResultField.sections.value] and \
        result[ResultField.abbrevs.value] and \
        result[ResultField.dictionary.value] and \
        result[ResultField.introoutro.value] and \
        result[ResultField.sources.value] and \
        result[ResultField.addition.value]

    result[ResultField.headers.value] = result[ResultField.contents.value]
    return result


def analyze(filename):
    contents = None
    with open(filename, mode='rb') as file:
        contents = file.read()

    if contents is None:
        raise Exception(DocExceptions.NOFILE)

    if not OLE.isOleFile(data=contents):
        raise Exception(DocExceptions.CORRUPT)

    with OLE.OleFileIO(contents) as ole:
        return process_ole(ole)
