#!/bin/bash

# 自动设置 CURRENT_DIR 为脚本所在目录（如果未设置）
if [ -z "$CURRENT_DIR" ]; then
    CURRENT_DIR=$(dirname "$(realpath "$0")")
    echo "CURRENT_DIR is automatically set to $CURRENT_DIR"
fi

# 遍历当前目录中的所有 .out 文件
for OUT_FILE in "$CURRENT_DIR"/*.out; do
    # 检查文件是否存在
    if [ ! -f "$OUT_FILE" ]; then
        echo "No .out files found in $CURRENT_DIR"
        continue
    fi

    OUT_FILE=$(basename "$OUT_FILE")
    OUT_SERVICE_FILE="/etc/systemd/system/${OUT_FILE}.service"

    # 删除现有的 .out 服务（如果存在）
    if [ -f "$OUT_SERVICE_FILE" ]; then
        echo "Stopping and removing existing service for $OUT_FILE"
        sudo systemctl stop "${OUT_FILE}.service"
        sudo systemctl disable "${OUT_FILE}.service"
        sudo rm -f "$OUT_SERVICE_FILE"
    fi

    sudo chmod +x "$CURRENT_DIR/$OUT_FILE"

    # 创建新的 .out systemd 服务
    echo "Creating systemd service for $OUT_FILE at $OUT_SERVICE_FILE"
    sudo bash -c "cat > $OUT_SERVICE_FILE" <<EOL
[Unit]
Description=Service for $OUT_FILE
After=network.target

[Service]
ExecStart=$CURRENT_DIR/$OUT_FILE
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOL

    # 启用并启动新的 .out 服务
    sudo systemctl daemon-reload
    sudo systemctl enable "${OUT_FILE}.service"
    sudo systemctl start "${OUT_FILE}.service"

    # 检查服务状态
    sudo systemctl status "${OUT_FILE}.service"
done