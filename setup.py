#!/usr/bin/env python
import pathlib
import os
from setuptools import setup
# from scriptforge_stubs import __version__
PKG_NAME = 'types-scriptforge'
VERSION = "1.1.2"

# The directory containing this file
HERE = pathlib.Path(__file__).parent
# The text of the README file
with open(HERE / "README.rst") as fh:
    README = fh.read()

src_path = str(HERE / 'scriptforge')

setup(
    name=PKG_NAME,
    version=VERSION,
    package_data={"": ["*.pyi", "py.typed"]},
    python_requires='>=3.7.0',
    url="https://github.com/Amourspirit/python-types-scriptforge",
    packages=["scriptforge"],
    author=":Barry-Thomas-Paul: Moss",
    author_email='bigbytetech@gmail.com',
    license="Apache Software License",
    keywords=['libreoffice', 'openoffice', 'scriptforge', 'typings', 'uno', 'ooouno', 'pyuno'],
    classifiers=[
        "License :: OSI Approved :: Apache Software License",
        "Environment :: Other Environment",
        "Intended Audience :: Developers",
        "Operating System :: OS Independent",
        "Topic :: Office/Business",
        "Typing :: Typed",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
    install_requires=[
        'typing_extensions>=3.7.4.3;python_version<"3.7"',
        'types-unopy>=0.1.7'
    ],
    description="Type annotations for ScriptForge",
    long_description_content_type="text/x-rst",
    long_description=README
)