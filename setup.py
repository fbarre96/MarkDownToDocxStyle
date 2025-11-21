from setuptools import find_packages, setup
import pathlib

# The directory containing this file
HERE = pathlib.Path(__file__).parent

# The text of the README file
README = (HERE / "README.md").read_text()

# Read requirements from requirements.txt with fallback
def get_requirements():
    """Read requirements from requirements.txt with fallback to hardcoded list."""
    requirements_file = HERE / "requirements.txt"
    
    # Fallback requirements in case requirements.txt is not available
    fallback_requirements = [
        "python-docx>=1.1.2",
        "requests",
        "pygments"
    ]
    
    try:
        with open(requirements_file, "r", encoding="utf-8") as f:
            reqs = [line.strip() for line in f if line.strip() and not line.startswith("#")]
        return reqs if reqs else fallback_requirements
    except FileNotFoundError:
        return fallback_requirements

requirements = get_requirements()

setup(
    name='markdowntodocx',
    packages=find_packages(include=['markdowntodocx']),
    version='0.1.8.1',
    url="https://github.com/fbarre96/MarkDownToDocxStyle",
    description='Convert markdown inside Docx to docx styles',
    long_description=README,
    long_description_content_type="text/markdown",
    author='Fabien Barre',
    license='MIT',
    install_requires=requirements,
)
