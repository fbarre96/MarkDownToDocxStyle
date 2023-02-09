from setuptools import find_packages, setup
import pathlib

# The directory containing this file
HERE = pathlib.Path(__file__).parent

# The text of the README file
README = (HERE / "README.md").read_text()

setup(
    name='markdowntodocx',
    packages=find_packages(include=['markdowntodocx']),
    version='0.1.3.3',
    url="https://github.com/fbarre96/MarkDownToDocxStyle",
    description='Convert markdown inside Docx to docx styles',
    long_description=README,
    long_description_content_type="text/markdown",
    author='Fabien Barre',
    license='MIT',
    setup_requires=['python-docx',"requests"],
)
