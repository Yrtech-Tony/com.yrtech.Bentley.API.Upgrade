using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class ExpenseAccountStatusCountDto
    {
       /* 1:申请提交材料-报价单
            2:申请提交材料-邮件截图
            3:报销证明材料-合同
            4:报销证明材料-发票
            5:报销证明材料-报价单
            6:报销证明材料-其他
            7:申请提交材料-活动计划
            8:报销证明材料-活动报告
            9:报销证明材料-邮件截图*/
        public Nullable<int> ExpenseAccount1 { get; set; }
        public Nullable<int> ExpenseAccount2 { get; set; }
        public Nullable<int> ExpenseAccount3 { get; set; }
        public Nullable<int> ExpenseAccount4 { get; set; }
        public Nullable<int> ExpenseAccount5 { get; set; }
        public Nullable<int> ExpenseAccount9 { get; set; }

    }
}