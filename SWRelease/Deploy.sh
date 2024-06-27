BIN="ems"

function copy() {
    path=$(dirname $2)
    if [ ! -d "${path}" ]; then
        echo "Create Dir:${path}"
        sudo mkdir -p "${path}"
    fi 
    echo "Copy $1 to $2"
    sudo cp -R "$1" "$2"
}

CURRENT_DIR=$(dirname $0)

# 拷贝可执行文件
copy $CURRENT_DIR/bin/x64/$BIN /usr/local/bin/$BIN

# 拷贝系统服务文件
copy $CURRENT_DIR/service/$BIN.service /etc/systemd/system/$BIN.service
#sudo cp service/happynet.service /etc/systemd/system/

sudo systemctl stop $BIN

# reload service
sudo systemctl daemon-reload

# start $ service when reboot
sudo systemctl enable $BIN

# start BIN service status
sudo systemctl start $BIN

# display BIN service status
sudo systemctl status $BIN

copy /usr/local/bin/*.a2l $CURRENT_DIR/*.a2l