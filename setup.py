from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="EWSModule",
    version="1.0.0.0",
    author="wns",
    author_email="wnsoff@yandex.ru",
    description="Wrap around exchangelib, for better expirience",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/wns444/EWSModule.git",
    packages=find_packages(),
    install_requires=[
        "exchangelib==5.5.0",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Topic :: Communications :: Email",
    ],
    python_requires=">=3.8",
    keywords="exchange ews email outlook",
)