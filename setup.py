from setuptools import setup, find_packages

setup(
    name="gspread-wrapper",
    version="0.1.0",
    author="Chris Buetti",
    author_email="crb4595@gmail.com",
    description="A wrapper for gspread with pandas support and robust error handling",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/chrisbuetti/gspread-wrapper",
    packages=find_packages(),
    install_requires=[
        "gspread>=5.0.0",
        "pandas>=1.0.0",
        "requests>=2.0.0"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: MIT License"
    ],
    python_requires='>=3.7',
)
