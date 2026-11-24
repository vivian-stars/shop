# 电商工具

按平台扩展的电商工具站。目前只有速卖通运费。

新平台在 `shop/shopapp/platforms/` 增加一个包，实现 `shop/shopapp/platforms/base.py` 里的接口，再在 `shop/shopapp/platforms/__init__.py` 注册。计算页面不用改。

## 运行

```bash
pip install -r requirements.txt
python -m shop.server
```

打开 http://127.0.0.1:3460 。

## 运费

配送服务费在物流费表里的单位是元/千克。产品运费 = 计费重量(kg) × 配送服务费 + 挂号费。表头是「0.5KG 首重 + 每 500G 续重」时，用首重运费加上超出首重的档数乘续重单价。产品最终运费再乘上六项成本比例，金额保留两位小数。

页面可以上传新的物流费表。上传之后，渠道列表和全部计算都改用新表。

`shop/data/default_rates.xlsx` 是一份工作摘录，国家少于完整的物流费表。要用全部国家的价格，请在页面上传完整的 `物流费.xlsx`。

## 测试

```bash
python -m pytest
```
