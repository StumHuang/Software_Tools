#!/bin/bash

# 解析参数
MODE="deploy"
while getopts "du" opt; do
    case $opt in
        d) MODE="deploy" ;;
        u) MODE="undeploy" ;;
        *) echo "Usage: $0 [-d|-u]"; exit 1 ;;
    esac
done

# 自动设置脚本目录
if [ -z "$CURRENT_DIR" ]; then
    CURRENT_DIR=$(dirname "$(realpath "$0")")
    echo "CURRENT_DIR is automatically set to $CURRENT_DIR"
fi

# 创建输出目录并设置权限
OUTPUT_DIR="$CURRENT_DIR/output"
if [ ! -d "$OUTPUT_DIR" ]; then
    mkdir -p "$OUTPUT_DIR"
    echo "Created output directory: $OUTPUT_DIR"
fi
sudo chmod 777 "$OUTPUT_DIR"

# 遍历所有 .out 文件
FOUND_OUT=0
for OUT_PATH in "$CURRENT_DIR"/*.out; do
    if [ ! -f "$OUT_PATH" ]; then
        continue
    fi
    FOUND_OUT=1
    OUT_FILE=$(basename "$OUT_PATH")
    OUT_SERVICE_FILE="/etc/systemd/system/${OUT_FILE}.service"

    if [ "$MODE" = "undeploy" ]; then
        # 仅移除部署
        if [ -f "$OUT_SERVICE_FILE" ]; then
            echo "Stopping and removing service for $OUT_FILE"
            sudo systemctl stop "${OUT_FILE}.service"
            sudo systemctl disable "${OUT_FILE}.service"
            sudo rm -f "$OUT_SERVICE_FILE"
            sudo systemctl daemon-reload
        else
            echo "Service file $OUT_SERVICE_FILE does not exist."
        fi
        continue
    fi

    # 部署流程
    # 停止并删除已有服务
    if [ -f "$OUT_SERVICE_FILE" ]; then
        echo "Stopping and removing existing service for $OUT_FILE"
        sudo systemctl stop "${OUT_FILE}.service"
        sudo systemctl disable "${OUT_FILE}.service"
        sudo rm -f "$OUT_SERVICE_FILE"
    fi

    sudo chmod +x "$OUT_PATH"

    # 创建 systemd 服务文件
    echo "Creating systemd service for $OUT_FILE at $OUT_SERVICE_FILE"
    sudo bash -c "cat > $OUT_SERVICE_FILE" <<EOL
[Unit]
Description=Service for $OUT_FILE
After=network.target

[Service]
WorkingDirectory=$CURRENT_DIR
ExecStart=$CURRENT_DIR/$OUT_FILE
Restart=always
RestartSec=5
# StandardOutput=append:$OUTPUT_DIR/${OUT_FILE}.log
# StandardError=append:$OUTPUT_DIR/${OUT_FILE}.err

[Install]
WantedBy=multi-user.target
EOL

    # 启用并启动服务
    sudo systemctl daemon-reload
    sudo systemctl enable "${OUT_FILE}.service"
    sudo systemctl start "${OUT_FILE}.service"
    sudo systemctl status "${OUT_FILE}.service" --no-pager
done

if [ $FOUND_OUT -eq 0 ]; then
    echo "No .out files found in $CURRENT_DIR"
fi