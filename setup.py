from setuptools import find_packages, setup

setup(
    name="headingdocx",
    version="0.1.0",
    description="A toolkit for manipulating Word documents by heading structure, including heading extraction, reordering, and XML operations.",
    author="YankoQiu",
    author_email="yanko.qiu@gmail.com",
    packages=find_packages(),
    install_requires=["python-docx>=0.8.11"],
    python_requires=">=3.6",
    url="",
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: MIT License",
        "Intended Audience :: Developers",
        "Topic :: Text Processing :: General",
    ],
    keywords="word docx heading outline xml python-docx",
)
