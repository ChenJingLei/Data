using System;
using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class Info
    {
        [Key]
        public int Id { get; set; }
        public string PROD_INST_ID { get; set; }
        public string 宽带账号 { get; set; }
        public string 关联ITV账号 { get; set; }
        public string 套餐 { get; set; }
        public string 支局 { get; set; }
        public string 网格 { get; set; }
        public string CUST_ID { get; set; }
        public string CUST_NAME { get; set; }
        public string AMOUNT { get; set; }
        public string EPON_TYPE { get; set; }
        public string 停机类型 { get; set; }
        public string 宽带速率属性ID { get; set; }
        public string 速率 { get; set; }
        public string 用户联系名称 { get; set; }
        public string 用户联系方式 { get; set; }
        public DateTime CRM竣工时间 { get; set; }
        public string 装机地址 { get; set; }
        public string 客户地址 { get; set; }
    }
}
