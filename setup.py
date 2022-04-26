import os
import setuptools
from setuptools import setup, find_packages

setup(
    name="googlesheet",
    version="0.0.1",
    description="a library for doing stuff with google sheets",
    python_requires=">=3.4",
    author="Adam Frank",
    author_email="pkgmaint@antilogo.org",
    packages=find_packages(),
    install_requires=["google-auth-oauthlib","google-api-python-client"],
    project_urls={"Source": "https://github.com/afrank/googlesheet",},
)
