from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="officekit",
    version="0.1.0",
    description="Excel / PPT / Word 자동화 유틸리티 (Win32 COM 미사용).",
    long_description=long_description,
    long_description_content_type="text/markdown",
    license="MIT",
    author="공튼이",
    author_email="wonseokie@hotmail.com",
    url="https://github.com/prozac0401/python_office",
    packages=find_packages(exclude=("tests", "docs")),
    include_package_data=True,
    install_requires=[
        "pandas>=2.0",
        "openpyxl>=3.1",
        "python-pptx>=0.6.23",
        "python-docx>=1.1",
        "matplotlib>=3.7",
    ],
    python_requires=">=3.9",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
    ],
)
