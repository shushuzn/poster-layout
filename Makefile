# 通用墙报排版 —— 常用任务
#
# 用法：  make            # 等价于 make build
#         make test
#         make clean
#
# Windows 未安装 make 时，直接执行等价命令（见 README）。

# 优先用仓库内的 .venv，没有才退回 PATH 上的 python
VENV_PY := $(firstword $(wildcard .venv/Scripts/python.exe) $(wildcard .venv/bin/python))
PYTHON  ?= $(if $(VENV_PY),$(VENV_PY),python)

BUILD  := $(PYTHON) build.py
TASK   ?=
DIST   := dist

.PHONY: help venv build check list new test test-all clean distclean setup lint fmt tree

help:  ## 显示本帮助
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) \
		| awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-12s\033[0m %s\n", $$1, $$2}'

venv:  ## 建虚拟环境并安装依赖
	python -m venv .venv
	$(PYTHON) -m pip install -q --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt

build:  ## 生成 PPTX + 预览图（TASK=<id> 可指定任务）
	$(BUILD) $(TASK)

check:  ## 构建 + 跑合规自检（清单见 docs/compliance.md）
	$(BUILD) --check $(TASK)

list:  ## 列出所有墙报任务
	$(BUILD) --list

new:  ## 新建任务：make new ID=my-poster
	@test -n "$(ID)" || (echo "用法: make new ID=<任务id>"; exit 2)
	$(BUILD) --new $(ID)

test:  ## 运行测试（无需 pytest）
	$(PYTHON) tests/runner.py

test-all:  ## 装 pytest 后跑同一批用例
	$(PYTHON) -m pytest -q

setup:  ## 安装依赖
	$(PYTHON) -m pip install -r requirements.txt

lint:  ## 静态检查（需 ruff）
	$(PYTHON) -m ruff check src build.py tests

fmt:  ## 代码格式化（需 ruff）
	$(PYTHON) -m ruff format src build.py tests

clean:  ## 清理可再生缓存与预览图
	rm -rf .cache
	rm -f $(DIST)/*.png

distclean: clean  ## 连交付物一起清理
	rm -f $(DIST)/*.pptx

tree:  ## 打印项目结构
	@find . -path ./.git -prune -o -path ./.venv -prune -o -path ./.cache -prune \
		-o -path ./assets/source -prune -o -type f -print | sort
