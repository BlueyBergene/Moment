from setuptools import setup, find_packages

setup(
    name="moment",
    version="0.2",
    packages=find_packages(),
    install_requires=[
        'click',
        'openpyxl'
    ],
    entry_points="""[console_scripts]
    moment=mover:main"""
)