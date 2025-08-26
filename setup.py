"""
银图PMC智能分析平台 - 安装配置
Yintu PMC Intelligent Analysis Platform - Setup Configuration
"""

from setuptools import setup, find_packages

with open("requirements.txt", encoding="utf-8") as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="yintu-pmc",
    version="1.0.0",
    author="Yintu PMC Team",
    description="银图PMC智能分析平台 - 生产计划与物料控制系统",
    long_description="A comprehensive production planning and material control system for hair dryer manufacturing",
    python_requires=">=3.8",
    packages=find_packages(),
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "yintu-dashboard=streamlit_dashboard:main",
            "yintu-supplier-analysis=精准供应商物料分析系统:main",
            "yintu-order-analysis=精准订单物料分析系统:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Manufacturing",
        "Topic :: Office/Business :: Financial :: Accounting",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
)