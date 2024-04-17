import os
import shutil

import setuptools

from expyvalidations import __version__

current_dir = os.getcwd()
dist_dir = f'{current_dir}/dist'
build_dir = f'{current_dir}/build'

if os.path.isdir(dist_dir):
    shutil.rmtree(dist_dir)

if os.path.isdir(build_dir):
    shutil.rmtree(build_dir)

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open('requirements.txt') as f:
    requirements = f.read().splitlines()

setuptools.setup(
    name="expyvalidations",
    packages=setuptools.find_packages(),
    version=__version__,
    description="Sheets data validation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="exebixel",
    author_email="ezequielnat7@gmail.com",
    url="https://github.com/exebixel/expyvalidations",
    license='MIT',
    install_requires=requirements,
    python_requires=">=3.7",
    classifiers=[
        "Operating System :: OS Independent",
        "Environment :: Console",
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Libraries',
        'License :: OSI Approved :: MIT License',
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
    ],
    project_urls={
        "Documentation": "https://github.com/exebixel/expyvalidations/wiki",
        "Source": "https://github.com/exebixel/expyvalidations",
    }
)
