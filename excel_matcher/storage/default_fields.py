from excel_matcher.models import StandardField


DEFAULT_FIELDS = [
    StandardField(
        key="order_no",
        display_name="订单号",
        domain="order",
        description="销售、采购或业务订单编号",
        aliases=["订单编号", "单号", "销售单号"],
    ),
    StandardField(
        key="customer_name",
        display_name="客户名称",
        domain="crm",
        description="客户、购买方或企业名称",
        aliases=["购买方", "企业名称", "客户名", "公司名称"],
    ),
    StandardField(
        key="amount",
        display_name="金额",
        domain="finance",
        description="订单、应收或含税金额",
        aliases=["订单金额", "含税金额", "应收金额", "总金额"],
    ),
    StandardField(
        key="order_date",
        display_name="下单日期",
        domain="order",
        description="订单或业务发生日期",
        aliases=["订单日期", "日期", "业务日期"],
    ),
    StandardField(
        key="product_name",
        display_name="商品名称",
        domain="inventory",
        description="商品、产品或物料名称",
        aliases=["产品名称", "物料名称", "品名"],
    ),
    StandardField(
        key="quantity",
        display_name="数量",
        domain="order",
        description="订购、采购或销售数量",
        aliases=["订购数量", "采购数量", "销售数量"],
    ),
    StandardField(
        key="sku",
        display_name="SKU",
        domain="inventory",
        description="商品、物料或产品编码",
        aliases=["商品编码", "物料编码", "产品编码"],
    ),
    StandardField(
        key="inventory_qty",
        display_name="库存数量",
        domain="inventory",
        description="现存或可用库存数量",
        aliases=["库存", "现存量", "可用库存"],
    ),
    StandardField(
        key="contact_name",
        display_name="联系人",
        domain="crm",
        description="客户联系人姓名",
        aliases=["联系人姓名", "客户联系人"],
    ),
    StandardField(
        key="phone",
        display_name="电话",
        domain="crm",
        description="客户联系电话或手机",
        aliases=["手机号", "联系电话", "手机"],
    ),
    StandardField(
        key="invoice_no",
        display_name="发票号",
        domain="finance",
        description="发票或票据号码",
        aliases=["发票号码", "票据号"],
    ),
    StandardField(
        key="paid_status",
        display_name="付款状态",
        domain="finance",
        description="订单或账款支付状态",
        aliases=["是否付款", "是否支付", "支付状态"],
    ),
    StandardField(
        key="account_name",
        display_name="账户名称",
        domain="finance",
        description="银行账户或开户名称",
        aliases=["开户名", "银行账户名"],
    ),
    StandardField(
        key="department",
        display_name="部门",
        domain="crm",
        description="业务、所属或经办部门",
        aliases=["所属部门", "业务部门"],
    ),
]


def get_default_fields() -> list[StandardField]:
    return [field.model_copy(deep=True) for field in DEFAULT_FIELDS]
