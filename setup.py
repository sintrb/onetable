from setuptools import setup
import os, io
from onetable import __version__

here = os.path.abspath(os.path.dirname(__file__))
README = io.open(os.path.join(here, 'README.md'), encoding='UTF-8').read()
CHANGES = io.open(os.path.join(here, 'CHANGES.md'), encoding='UTF-8').read()
setup(name="onetable",
      version=__version__,
      keywords=('onetable', 'excel', 'csv', 'table', 'export'),
      description="Once defined, multiple ways to export table.",
      long_description=README + '\n\n\n' + CHANGES,
      long_description_content_type="text/markdown",
      url='https://github.com/sintrb/onetable/',
      author="trb",
      author_email="sintrb@gmail.com",
      packages=['onetable'],
      install_requires=[],
      zip_safe=False
      )
