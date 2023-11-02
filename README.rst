===================
ScriptForge Typings
===================

This project is Type Support (typings) for LibreOffice `ScriptForge <https://gitlab.com/LibreOfficiant/scriptforge>`_

At the time of this writing there is no other PyPi package for ScriptForge.
It is not necessary to have a ScriptForge package to take advantage of
ScriptForge Typings.

ScriptForge lives inside of LibreOffice. Using these typings in a modern development IDE
will give type Support for the ScriptForge library.

This project leverages `types-unopy <https://github.com/Amourspirit/python-types-unopy>`_ that gives
full typing support for `LibreOffice API <https://api.libreoffice.org/>`_.
This allows full type support for `ScriptForge <https://gitlab.com/LibreOfficiant/scriptforge>`_
and `LibreOffice API <https://api.libreoffice.org/>`_.

These Typings are created for LibreOffice ``7.6``.

Installation
============

PIP
---

**types-scriptforge** `PyPI <https://pypi.org/project/types-scriptforge/>`_

.. code-block:: bash

    $ pip install types-scriptforge

`ScriptForge Docs <https://help.libreoffice.org/latest/en-US/text/sbasic/shared/03/lib_ScriptForge.html>`__ on LibreOffice Help.

Older versions
--------------

While this version will also work with previous version of ScriptForge, not all methods/functions are available in previous versions
that are part of this typing library.

To install for version ``7.5``.

.. code-block:: bash

    $ pip install "types-scriptforge>=2.0,<3.0"

To install for version ``7.4``.

.. code-block:: bash

    $ pip install "types-scriptforge>=1.1,<2.0"

To install for version ``7.3``.

.. code-block:: bash

    $ pip install "types-scriptforge<1.1"

Other
=====

**Figure 1:** ScriptForge typings example

.. figure:: https://user-images.githubusercontent.com/4193389/163497042-a572dff9-0278-4d42-be22-dea4555545ff.gif
   :alt: types-scriptforge example gif.