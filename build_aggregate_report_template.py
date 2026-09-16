# -*- coding: utf-8 -*-
"""Tạo lại template Word giải trình tổng hợp mặc định."""

from src.aggregate_report_exporter import DEFAULT_TEMPLATE_PATH, build_default_template


if __name__ == "__main__":
    print(build_default_template(DEFAULT_TEMPLATE_PATH))
