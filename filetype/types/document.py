# -*- coding: utf-8 -*-

from __future__ import absolute_import

from .base import Type
from struct import unpack


class Doc(Type):
    """
    Implements the Microsoft Word Doc
    """
    MIME = 'application/msword'
    EXTENSION = 'doc'

    def __init__(self):
        super(Doc, self).__init__(
            mime=Doc.MIME,
            extension=Doc.EXTENSION
        )

    def match(self, buf):
        return (len(buf) > 7 and
                buf[0] == 0xD0 and buf[1] == 0xCF and
                buf[2] == 0x11 and buf[3] == 0xE0 and
                buf[4] == 0xA1 and buf[5] == 0xB1 and
                buf[6] == 0x1A and buf[7] == 0xE1)


class Xls(Type):
    """
    Implements the Microsoft Excel Xls
    """
    MIME = 'application/vnd.ms-excel'
    EXTENSION = 'xls'

    def __init__(self):
        super(Xls, self).__init__(
            mime=Xls.MIME,
            extension=Xls.EXTENSION
        )

    def match(self, buf):
        return (len(buf) > 7 and
                buf[0] == 0xD0 and buf[1] == 0xCF and
                buf[2] == 0x11 and buf[3] == 0xE0 and
                buf[4] == 0xA1 and buf[5] == 0xB1 and
                buf[6] == 0x1A and buf[7] == 0xE1)

class Ppt(Type):
    """
    Implements the Microsoft Powerpoint Ppt
    """
    MIME = 'application/vnd.ms-powerpoint'
    EXTENSION = 'ppt'

    def __init__(self):
        super(Ppt, self).__init__(
            mime=Ppt.MIME,
            extension=Ppt.EXTENSION
        )

    def match(self, buf):
        return (len(buf) > 7 and
                buf[0] == 0xD0 and buf[1] == 0xCF and
                buf[2] == 0x11 and buf[3] == 0xE0 and
                buf[4] == 0xA1 and buf[5] == 0xB1 and
                buf[6] == 0x1A and buf[7] == 0xE1)


class Msooxml(Type):
    """
    Implement msooxml standar
    """

    signature = bytearray([0x50,0x4B,0x03,0x04])

    def isMsooxml(self, buf):

        if not(self.compareBytes(buf, self.signature, 0)):
            return None, False

        v, ok = self.checkMSOoml(buf, 0x1E)
        if ok:
            return v, ok

        if not(self.compareBytes(buf, self.getByteArray('[Content_Types].xml'), 0x1E)) and not(self.compareBytes(buf, self.getByteArray('_rels/.rels'), 0x1E)):
            return 'content', False

        hi, lo = unpack('<hh', buf[18:22])
        #startOffset = ((hi << 16) | lo) + 49
        startOffset = hi + lo + 49
        idx = self.search(buf, startOffset, 6000)
        if idx == -1:
            return None, False

        startOffset += idx + 4 + 26
        idx = self.search(buf, startOffset, 6000)
        if idx == -1:
            return None, False

        startOffset += idx + 4 + 26
        v, ok = self.checkMSOoml(buf, startOffset)
        if ok:
            return v, ok

        # Openoffice
        startOffset += 26
        idx = self.search(buf, startOffset, 6000)
        if idx == -1:
            return 'TYPE_OOXML', True

        startOffset += idx + 4 + 26
        v, ok = self.checkMSOoml(buf, startOffset)
        if ok:
            return v, ok
        else:
            return 'TYPE_OOXML', True
        return None

    def compareBytes(self, slice, subSlice, startOffset):
        sl = len(subSlice)
        if startOffset + sl > len(slice):
            return False
        s = slice[startOffset:startOffset+sl]
        return s == subSlice

    def getByteArray(self, label):
        typedoc = bytearray()
        typedoc.extend(map(ord, label))
        return typedoc

    def checkMSOoml(self, buf, offset):
        ok = True
        if self.compareBytes(buf, self.getByteArray('word/'), offset):
            return 'type_docx', ok
        if self.compareBytes(buf, self.getByteArray('ppt/'), offset):
            return 'type_pptx', ok
        if self.compareBytes(buf, self.getByteArray('xl/'), offset):
            return 'type_xslx', ok
        ok = False
        return None, ok

    def search(self, buf, start, rangeNum):
        length = len(buf)
        end = start + rangeNum
        if end > length:
            end = length
        if start >= end:
            return -1
        return buf[start:end].find(self.signature)


class Docx(Msooxml):
    """
    Implements the Microsoft Word Docx
    """
    MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    EXTENSION = 'docx'

    def __init__(self):
        super(Docx, self).__init__(
            mime=Docx.MIME,
            extension=Docx.EXTENSION
        )

    def match(self, buf):
        typ, ok = self.isMsooxml(buf)
        return ok and typ == 'type_docx'

class Xslx(Msooxml):
    """
    Implements the Microsoft Word Docx
    """
    MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    EXTENSION = 'xslx'

    def __init__(self):
        super(Xslx, self).__init__(
            mime=Xslx.MIME,
            extension=Xslx.EXTENSION
        )

    def match(self, buf):
        typ, ok = self.isMsooxml(buf)
        return ok and typ == 'type_xslx'

class Pptx(Msooxml):
    """
    Implements the Microsoft Word Docx
    """
    MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    EXTENSION = 'pptx'

    def __init__(self):
        super(Pptx, self).__init__(
            mime=Pptx.MIME,
            extension=Pptx.EXTENSION
        )

    def match(self, buf):
        typ, ok = self.isMsooxml(buf)
        return ok and typ == 'type_pptx'
