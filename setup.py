#!/usr/bin/env python3
"""Setup script for mohanxlsx."""

from setuptools import setup, find_packages

setup(
    name="mohanxlsx",
    version="0.1.2",
    description="xlsx表格工具箱，支持拆分、合并、去除空行等",
    author="sd44",
    author_email="sd44sd44@yeah.net",
    license="GPL3",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.6",
    install_requires=[
        "openpyxl>=3.1.5",
        "pandas>=2.0",
        "qtpy>=2.4",
    ],
    entry_points={
        "console_scripts": [
            "mohanxlsx=mohanxlsx.app:main",
        ],
    },
    include_package_data=True,
)
