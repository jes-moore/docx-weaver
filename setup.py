"""
Setup for word-weaver
"""

from setuptools import setup, find_packages

with open("requirements.txt", encoding="utf-8") as f:
    requirements = f.read().splitlines()
setup(
    name="wordweaver",
    version="0.1.0",
    url="https://github.com/jes-moore/word-weaver",
    author="Jesse Moore",
    author_email="jessemoore07@gmail.com",
    description=(
        "WordWeaver is a Python class designed to convert, translate, "
        "and modify Word documents in various ways, such as adding comments, "
        "transforming text, and both. It's suitable for automating document "
        "workflows where changes to the document based on dynamic inputs are required."
    ),
    packages=find_packages(),
    install_requires=requirements,
)
