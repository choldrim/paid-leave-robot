# paid-leave-robot
此项目有毒，勿动

### 1. 运行`user.py`拉取所有用户的邮箱（存在的会跳过）每一月重新拉取一遍全新的
### 2. 放入overtime.xlsx(加班数据) leave.xlsx(请假数据)文件到对应月份的目录中(如计算2016年5月份的，就放到 `data/201605` 目录中)
### 3. 修改`config.ini`中的`month`字段(如计算的是2016年6月份的，month就为`2016-06`)
### 4. 运行./main.py --mail (--mail是运行并发送邮件)

