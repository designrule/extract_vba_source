#!/usr/bin/env python3

from shutil import rmtree
from pathlib import Path
from argparse import ArgumentParser
from oletools.olevba import VBA_Parser, VBA_Project, filter_vba

OFFICE_FILE_EXTENSIONS = (
    '.xlsb', '.xls', '.xlsm', '.xla', '.xlt', '.xlam',  # Excel book with macro
    '.pptm',
)


def get_args():
    parser = ArgumentParser(description='Extract vba source files from an MS Office file with macro.')
    parser.add_argument('sources', metavar='MS_OFFICE_FILE', type=str, nargs='+',
                        help='Paths to source MS Office file or directory.')
    parser.add_argument('--dest', type=str, default='vba_src',
                        help='Destination directory path to output vba source files [default: ./vba_src].')
    parser.add_argument('--orig-extension', dest='use_orig_extension', action='store_true',
                        help='Use an original extension (.bas, .cls, .frm) for extracted vba source files [default: use .vb].')
    parser.add_argument('--src-encoding', dest='src_encoding', type=str, default='shift_jis',
                        help='Encoding for vba source files in an MS Office file [default: shift_jis].')
    parser.add_argument('--out-encoding', dest='out_encoding', type=str, default='utf8',
                        help='Encoding for generated vba source files [default: utf8].')
    parser.add_argument('--recursive', action='store_true',
                        help='Find sub directories recursively when a directory is specified as the sources parameter.')
    return parser.parse_args()


def get_source_paths(sources, recursive):
    for src in sources:
        p = Path(src)
        if p.is_dir():  # If source is a directory, then find source files under it.
            for file in p.glob("**/*" if recursive else "*"):
                f = Path(file)
                if not f.name.startswith('~$') and f.suffix.lower() in OFFICE_FILE_EXTENSIONS:
                    yield f.absolute()
        else:  # If source is a file, then return its absolute path.
            yield p.absolute()


def get_outputpath(parent_dir: Path, filename: str, use_orig_extension: bool):
    extension = filename.split('.')[-1]
    if extension == 'cls':
        subdir = parent_dir.joinpath('class')
    elif extension == 'frm':
        subdir = parent_dir.joinpath('form')
    else:
        subdir = parent_dir.joinpath('module')

    if not subdir.exists():
        subdir.mkdir(parents=True, exist_ok=True)
    return Path(subdir.joinpath(filename + ('.vb' if not use_orig_extension else '')))

import traceback
import sys
import os
import struct
from io import BytesIO, StringIO
import math
import zipfile
import re
import argparse
import binascii
import base64
import zlib
import email  # for MHTML parsing
import email.feedparser
import string  # for printable
import json   # for json output mode (argument --json)

import olefile
from oletools.thirdparty.tablestream import tablestream
from oletools.thirdparty.xglob import xglob, PathNotFoundException
from oletools.thirdparty.oledump.plugin_biff import cBIFF
from oletools import ppt_parser
from oletools import oleform
from oletools import rtfobj
from oletools import crypto
from oletools.common.io_encoding import ensure_stdout_handles_unicode
from oletools.common import codepages
from oletools import ftguess
from oletools.common.log_helper import log_helper

# a global logger object used for debugging:
log = log_helper.get_or_create_silent_logger('olevba')

def vba_project_init(self, ole, vba_root, project_path, dir_path, relaxed=True):
    """
        Extract VBA macros from an OleFileIO object.
        :param vba_root: path to the VBA root storage, containing the VBA storage and the PROJECT stream
        :param project_path: path to the PROJECT stream
        :param relaxed: If True, only create info/debug log entry if data is not as expected
                        (e.g. opening substream fails); if False, raise an error in this case
        """
    self.ole = ole
    self.vba_root = vba_root
    self. project_path = project_path
    self.dir_path = dir_path
    self.relaxed = relaxed
    #: VBA modules contained in the project (list of VBA_Module objects)
    self.modules = []
    #: file extension for each VBA module
    self.module_ext = {}
    log.debug('Parsing the dir stream from %r' % dir_path)
    # read data from dir stream (compressed)
    dir_compressed = ole.openstream(dir_path).read()
    # decompress it:
    dir_stream = BytesIO(decompress_stream(bytearray(dir_compressed)))
    # store reference for later use:
    self.dir_stream = dir_stream

    # reference: MS-VBAL 2.3.4.2 dir Stream: Version Independent Project Information

    # PROJECTSYSKIND Record
    # Specifies the platform for which the VBA project is created.
    projectsyskind_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTSYSKIND_Id', 0x0001, projectsyskind_id)
    projectsyskind_size = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTSYSKIND_Size', 0x0004, projectsyskind_size)
    self.syskind = struct.unpack("<L", dir_stream.read(4))[0]
    SYSKIND_NAME = {
        0x00: "16-bit Windows",
        0x01: "32-bit Windows",
        0x02: "Macintosh",
        0x03: "64-bit Windows"
    }
    self.syskind_name = SYSKIND_NAME.get(self.syskind, 'Unknown')
    log.debug("PROJECTSYSKIND_SysKind: %d - %s" % (self.syskind, self.syskind_name))
    if self.syskind not in SYSKIND_NAME:
        log.error("invalid PROJECTSYSKIND_SysKind {0:04X}".format(self.syskind))

    # https://github.com/decalage2/oletools/issues/700#issuecomment-892562826
    # Temp Fix
    id_temp = struct.unpack("<H", dir_stream.read(2))[0]
    if id_temp == 0x004A:
        size_temp = struct.unpack("<L", dir_stream.read(4))[0]
        value_temp = struct.unpack("<L", dir_stream.read(size_temp))[0]
        id_temp = struct.unpack("<H", dir_stream.read(2))[0]

    # PROJECTLCID Record
    # Specifies the VBA project's LCID.
    projectlcid_id = id_tmp
    self.check_value('PROJECTLCID_Id', 0x0002, projectlcid_id)
    projectlcid_size = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTLCID_Size', 0x0004, projectlcid_size)
    # Lcid (4 bytes): An unsigned integer that specifies the LCID value for the VBA project. MUST be 0x00000409.
    self.lcid = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTLCID_Lcid', 0x409, self.lcid)

    # PROJECTLCIDINVOKE Record
    # Specifies an LCID value used for Invoke calls on an Automation server as specified in [MS-OAUT] section 3.1.4.4.
    projectlcidinvoke_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTLCIDINVOKE_Id', 0x0014, projectlcidinvoke_id)
    projectlcidinvoke_size = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTLCIDINVOKE_Size', 0x0004, projectlcidinvoke_size)
    # LcidInvoke (4 bytes): An unsigned integer that specifies the LCID value used for Invoke calls. MUST be 0x00000409.
    self.lcidinvoke = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTLCIDINVOKE_LcidInvoke', 0x409, self.lcidinvoke)

    # PROJECTCODEPAGE Record
    # Specifies the VBA project's code page.
    projectcodepage_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTCODEPAGE_Id', 0x0003, projectcodepage_id)
    projectcodepage_size = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTCODEPAGE_Size', 0x0002, projectcodepage_size)
    self.codepage = struct.unpack("<H", dir_stream.read(2))[0]
    self.codepage_name = codepages.get_codepage_name(self.codepage)
    log.debug('Project Code Page: %r - %s' % (self.codepage, self.codepage_name))
    self.codec = codepages.codepage2codec(self.codepage)
    log.debug('Python codec corresponding to code page %d: %s' % (self.codepage, self.codec))


    # PROJECTNAME Record
    # Specifies a unique VBA identifier as the name of the VBA project.
    projectname_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTNAME_Id', 0x0004, projectname_id)
    sizeof_projectname = struct.unpack("<L", dir_stream.read(4))[0]
    log.debug('Project name size: %d bytes' % sizeof_projectname)
    if sizeof_projectname < 1 or sizeof_projectname > 128:
        # TODO: raise an actual error? What is MS Office's behaviour?
        log.error("PROJECTNAME_SizeOfProjectName value not in range [1-128]: {0}".format(sizeof_projectname))
        projectname_bytes = dir_stream.read(sizeof_projectname)
        self.projectname = self.decode_bytes(projectname_bytes)


    # PROJECTDOCSTRING Record
    # Specifies the description for the VBA project.
    projectdocstring_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTDOCSTRING_Id', 0x0005, projectdocstring_id)
    projectdocstring_sizeof_docstring = struct.unpack("<L", dir_stream.read(4))[0]
    if projectdocstring_sizeof_docstring > 2000:
        log.error(
            "PROJECTDOCSTRING_SizeOfDocString value not in range: {0}".format(projectdocstring_sizeof_docstring))
        # DocString (variable): An array of SizeOfDocString bytes that specifies the description for the VBA project.
        # MUST contain MBCS characters encoded using the code page specified in PROJECTCODEPAGE (section 2.3.4.2.1.4).
        # MUST NOT contain null characters.
    docstring_bytes = dir_stream.read(projectdocstring_sizeof_docstring)
    self.docstring = self.decode_bytes(docstring_bytes)
    projectdocstring_reserved = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTDOCSTRING_Reserved', 0x0040, projectdocstring_reserved)
    projectdocstring_sizeof_docstring_unicode = struct.unpack("<L", dir_stream.read(4))[0]
    if projectdocstring_sizeof_docstring_unicode % 2 != 0:
        log.error("PROJECTDOCSTRING_SizeOfDocStringUnicode is not even")
    # DocStringUnicode (variable): An array of SizeOfDocStringUnicode bytes that specifies the description for the
    # VBA project. MUST contain UTF-16 characters. MUST NOT contain null characters.
    # MUST contain the UTF-16 encoding of DocString.
    docstring_unicode_bytes = dir_stream.read(projectdocstring_sizeof_docstring_unicode)
    self.docstring_unicode = docstring_unicode_bytes.decode('utf16', errors='replace')

    # PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
    projecthelpfilepath_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTHELPFILEPATH_Id', 0x0006, projecthelpfilepath_id)
    projecthelpfilepath_sizeof_helpfile1 = struct.unpack("<L", dir_stream.read(4))[0]
    if projecthelpfilepath_sizeof_helpfile1 > 260:
        log.error(
            "PROJECTHELPFILEPATH_SizeOfHelpFile1 value not in range: {0}".format(projecthelpfilepath_sizeof_helpfile1))
    projecthelpfilepath_helpfile1 = dir_stream.read(projecthelpfilepath_sizeof_helpfile1)
    projecthelpfilepath_reserved = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTHELPFILEPATH_Reserved', 0x003D, projecthelpfilepath_reserved)
    projecthelpfilepath_sizeof_helpfile2 = struct.unpack("<L", dir_stream.read(4))[0]
    if projecthelpfilepath_sizeof_helpfile2 != projecthelpfilepath_sizeof_helpfile1:
        log.error("PROJECTHELPFILEPATH_SizeOfHelpFile1 does not equal PROJECTHELPFILEPATH_SizeOfHelpFile2")
    projecthelpfilepath_helpfile2 = dir_stream.read(projecthelpfilepath_sizeof_helpfile2)
    if projecthelpfilepath_helpfile2 != projecthelpfilepath_helpfile1:
        log.error("PROJECTHELPFILEPATH_HelpFile1 does not equal PROJECTHELPFILEPATH_HelpFile2")

    # PROJECTHELPCONTEXT Record
    projecthelpcontext_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTHELPCONTEXT_Id', 0x0007, projecthelpcontext_id)
    projecthelpcontext_size = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTHELPCONTEXT_Size', 0x0004, projecthelpcontext_size)
    projecthelpcontext_helpcontext = struct.unpack("<L", dir_stream.read(4))[0]
    unused = projecthelpcontext_helpcontext

    # PROJECTLIBFLAGS Record
    projectlibflags_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTLIBFLAGS_Id', 0x0008, projectlibflags_id)
    projectlibflags_size = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTLIBFLAGS_Size', 0x0004, projectlibflags_size)
    projectlibflags_projectlibflags = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTLIBFLAGS_ProjectLibFlags', 0x0000, projectlibflags_projectlibflags)

    # PROJECTVERSION Record
    projectversion_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTVERSION_Id', 0x0009, projectversion_id)
    projectversion_reserved = struct.unpack("<L", dir_stream.read(4))[0]
    self.check_value('PROJECTVERSION_Reserved', 0x0004, projectversion_reserved)
    projectversion_versionmajor = struct.unpack("<L", dir_stream.read(4))[0]
    projectversion_versionminor = struct.unpack("<H", dir_stream.read(2))[0]
    unused = projectversion_versionmajor
    unused = projectversion_versionminor

    # PROJECTCONSTANTS Record
    projectconstants_id = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTCONSTANTS_Id', 0x000C, projectconstants_id)
    projectconstants_sizeof_constants = struct.unpack("<L", dir_stream.read(4))[0]
    if projectconstants_sizeof_constants > 1015:
        log.error(
            "PROJECTCONSTANTS_SizeOfConstants value not in range: {0}".format(projectconstants_sizeof_constants))
    projectconstants_constants = dir_stream.read(projectconstants_sizeof_constants)
    projectconstants_reserved = struct.unpack("<H", dir_stream.read(2))[0]
    self.check_value('PROJECTCONSTANTS_Reserved', 0x003C, projectconstants_reserved)
    projectconstants_sizeof_constants_unicode = struct.unpack("<L", dir_stream.read(4))[0]
    if projectconstants_sizeof_constants_unicode % 2 != 0:
        log.error("PROJECTCONSTANTS_SizeOfConstantsUnicode is not even")
    projectconstants_constants_unicode = dir_stream.read(projectconstants_sizeof_constants_unicode)
    unused = projectconstants_constants
    unused = projectconstants_constants_unicode

    # array of REFERENCE records
    # Specifies a reference to an Automation type library or VBA project.
    check = None
    while True:
        check = struct.unpack("<H", dir_stream.read(2))[0]
        log.debug("reference type = {0:04X}".format(check))
        if check == 0x000F:
            break

        if check == 0x0016:
            # REFERENCENAME
            # Specifies the name of a referenced VBA project or Automation type library.
            reference_id = check
            reference_sizeof_name = struct.unpack("<L", dir_stream.read(4))[0]
            reference_name = dir_stream.read(reference_sizeof_name)
            log.debug('REFERENCE name: %s' % unicode2str(self.decode_bytes(reference_name)))
            reference_reserved = struct.unpack("<H", dir_stream.read(2))[0]
            # According to [MS-OVBA] 2.3.4.2.2.2 REFERENCENAME Record:
            # "Reserved (2 bytes): MUST be 0x003E. MUST be ignored."
            # So let's ignore it, otherwise it crashes on some files (issue #132)
            # PR #135 by @c1fe:
            # contrary to the specification I think that the unicode name
            # is optional. if reference_reserved is not 0x003E I think it
            # is actually the start of another REFERENCE record
            # at least when projectsyskind_syskind == 0x02 (Macintosh)
            if reference_reserved == 0x003E:
                #if reference_reserved not in (0x003E, 0x000D):
                #    raise UnexpectedDataError(dir_path, 'REFERENCE_Reserved',
                #                              0x0003E, reference_reserved)
                reference_sizeof_name_unicode = struct.unpack("<L", dir_stream.read(4))[0]
                reference_name_unicode = dir_stream.read(reference_sizeof_name_unicode)
                unused = reference_id
                unused = reference_name
                unused = reference_name_unicode
                continue
            else:
                check = reference_reserved
                log.debug("reference type = {0:04X}".format(check))

        if check == 0x0033:
            # REFERENCEORIGINAL (followed by REFERENCECONTROL)
            # Specifies the identifier of the Automation type library the containing REFERENCECONTROL's
            # (section 2.3.4.2.2.3) twiddled type library was generated from.
            referenceoriginal_id = check
            referenceoriginal_sizeof_libidoriginal = struct.unpack("<L", dir_stream.read(4))[0]
            referenceoriginal_libidoriginal = dir_stream.read(referenceoriginal_sizeof_libidoriginal)
            log.debug('REFERENCE original lib id: %s' % unicode2str(self.decode_bytes(referenceoriginal_libidoriginal)))
            unused = referenceoriginal_id
            unused = referenceoriginal_libidoriginal
            continue

        if check == 0x002F:
            # REFERENCECONTROL
            # Specifies a reference to a twiddled type library and its extended type library.
            referencecontrol_id = check
            referencecontrol_sizetwiddled = struct.unpack("<L", dir_stream.read(4))[0]  # ignore
            referencecontrol_sizeof_libidtwiddled = struct.unpack("<L", dir_stream.read(4))[0]
            referencecontrol_libidtwiddled = dir_stream.read(referencecontrol_sizeof_libidtwiddled)
            log.debug('REFERENCE control twiddled lib id: %s' % unicode2str(self.decode_bytes(referencecontrol_libidtwiddled)))
            referencecontrol_reserved1 = struct.unpack("<L", dir_stream.read(4))[0]  # ignore
            self.check_value('REFERENCECONTROL_Reserved1', 0x0000, referencecontrol_reserved1)
            referencecontrol_reserved2 = struct.unpack("<H", dir_stream.read(2))[0]  # ignore
            self.check_value('REFERENCECONTROL_Reserved2', 0x0000, referencecontrol_reserved2)
            unused = referencecontrol_id
            unused = referencecontrol_sizetwiddled
            unused = referencecontrol_libidtwiddled
            # optional field
            check2 = struct.unpack("<H", dir_stream.read(2))[0]
            if check2 == 0x0016:
                referencecontrol_namerecordextended_id = check
                referencecontrol_namerecordextended_sizeof_name = struct.unpack("<L", dir_stream.read(4))[0]
                referencecontrol_namerecordextended_name = dir_stream.read(
                    referencecontrol_namerecordextended_sizeof_name)
                log.debug('REFERENCE control name record extended: %s' % unicode2str(
                    self.decode_bytes(referencecontrol_namerecordextended_name)))
                referencecontrol_namerecordextended_reserved = struct.unpack("<H", dir_stream.read(2))[0]
                if referencecontrol_namerecordextended_reserved == 0x003E:
                    referencecontrol_namerecordextended_sizeof_name_unicode = struct.unpack("<L", dir_stream.read(4))[0]
                    referencecontrol_namerecordextended_name_unicode = dir_stream.read(
                        referencecontrol_namerecordextended_sizeof_name_unicode)
                    referencecontrol_reserved3 = struct.unpack("<H", dir_stream.read(2))[0]
                    unused = referencecontrol_namerecordextended_id
                    unused = referencecontrol_namerecordextended_name
                    unused = referencecontrol_namerecordextended_name_unicode
                else:
                    referencecontrol_reserved3 = referencecontrol_namerecordextended_reserved
            else:
                referencecontrol_reserved3 = check2

            self.check_value('REFERENCECONTROL_Reserved3', 0x0030, referencecontrol_reserved3)
            referencecontrol_sizeextended = struct.unpack("<L", dir_stream.read(4))[0]
            referencecontrol_sizeof_libidextended = struct.unpack("<L", dir_stream.read(4))[0]
            referencecontrol_libidextended = dir_stream.read(referencecontrol_sizeof_libidextended)
            referencecontrol_reserved4 = struct.unpack("<L", dir_stream.read(4))[0]
            referencecontrol_reserved5 = struct.unpack("<H", dir_stream.read(2))[0]
            referencecontrol_originaltypelib = dir_stream.read(16)
            referencecontrol_cookie = struct.unpack("<L", dir_stream.read(4))[0]
            unused = referencecontrol_sizeextended
            unused = referencecontrol_libidextended
            unused = referencecontrol_reserved4
            unused = referencecontrol_reserved5
            unused = referencecontrol_originaltypelib
            unused = referencecontrol_cookie
            continue

        if check == 0x000D:
            # REFERENCEREGISTERED
            # Specifies a reference to an Automation type library.
            referenceregistered_id = check
            referenceregistered_size = struct.unpack("<L", dir_stream.read(4))[0]
            referenceregistered_sizeof_libid = struct.unpack("<L", dir_stream.read(4))[0]
            referenceregistered_libid = dir_stream.read(referenceregistered_sizeof_libid)
            log.debug('REFERENCE registered lib id: %s' % unicode2str(self.decode_bytes(referenceregistered_libid)))
            referenceregistered_reserved1 = struct.unpack("<L", dir_stream.read(4))[0]
            self.check_value('REFERENCEREGISTERED_Reserved1', 0x0000, referenceregistered_reserved1)
            referenceregistered_reserved2 = struct.unpack("<H", dir_stream.read(2))[0]
            self.check_value('REFERENCEREGISTERED_Reserved2', 0x0000, referenceregistered_reserved2)
            unused = referenceregistered_id
            unused = referenceregistered_size
            unused = referenceregistered_libid
            continue

        if check == 0x000E:
            # REFERENCEPROJECT
            # Specifies a reference to an external VBA project.
            referenceproject_id = check
            referenceproject_size = struct.unpack("<L", dir_stream.read(4))[0]
            referenceproject_sizeof_libidabsolute = struct.unpack("<L", dir_stream.read(4))[0]
            referenceproject_libidabsolute = dir_stream.read(referenceproject_sizeof_libidabsolute)
            log.debug('REFERENCE project lib id absolute: %s' % unicode2str(self.decode_bytes(referenceproject_libidabsolute)))
            referenceproject_sizeof_libidrelative = struct.unpack("<L", dir_stream.read(4))[0]
            referenceproject_libidrelative = dir_stream.read(referenceproject_sizeof_libidrelative)
            log.debug('REFERENCE project lib id relative: %s' % unicode2str(self.decode_bytes(referenceproject_libidrelative)))
            referenceproject_majorversion = struct.unpack("<L", dir_stream.read(4))[0]
            referenceproject_minorversion = struct.unpack("<H", dir_stream.read(2))[0]
            unused = referenceproject_id
            unused = referenceproject_size
            unused = referenceproject_libidabsolute
            unused = referenceproject_libidrelative
            unused = referenceproject_majorversion
            unused = referenceproject_minorversion
            continue

        log.error('invalid or unknown check Id {0:04X}'.format(check))
        # raise an exception instead of stopping abruptly (issue #180)
        raise UnexpectedDataError(dir_path, 'reference type', (0x0F, 0x16, 0x33, 0x2F, 0x0D, 0x0E), check)
        #sys.exit(0)

def extract_macros(parser: VBA_Parser, vba_encoding):

    if parser.ole_file is None:
        for subfile in parser.ole_subfiles:
            for results in extract_macros(subfile, vba_encoding):
                yield results
    else:
        VBA_Project.__init__ = vba_project_init
        parser.find_vba_projects()
        for (vba_root, project_path, dir_path) in parser.vba_projects:
            project = VBA_Project(parser.ole_file, vba_root, project_path, dir_path, relaxed=True)
            project.codec = vba_encoding
            project.parse_project_stream()

            for code_path, vba_filename, code_data in project.parse_modules():
                yield (vba_filename, code_data)

def main():
    args = get_args()

    # Get the root path of destination (if not exists then make it).
    root = Path(args.dest)
    if not root.exists():
        root.mkdir(parents=True)
    elif not root.is_dir():
        raise FileExistsError

    # Get the source MS Office file where extract the vba source files from.
    for source in get_source_paths(args.sources, args.recursive):
        src = Path(source)
        basename = src.stem
        dest = Path(root.joinpath(basename))
        dest.mkdir(parents=True, exist_ok=True)
        rmtree(dest.absolute())
        print('Extract vba files from {source} to {dest}'.format(source=source, dest=dest))

        # Extract vba source files from the MS Office file and save each vba file into the sub directory as of its MS Office file name.
        vba_parser = VBA_Parser(src)
        for vba_filename, vba_code in extract_macros(vba_parser, args.src_encoding):
            vba_file = get_outputpath(dest, vba_filename, args.use_orig_extension)
            vba_file.write_text(filter_vba(vba_code), encoding=args.out_encoding)
            print('[{basename}] {vba_file} is generated.'.format(basename=basename, vba_file=vba_file))



if __name__ == '__main__':
    main()
