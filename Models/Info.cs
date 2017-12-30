using System;

namespace Data.Models
{
    [Key]
    public	int	Id	{get;set;}
    public	float	PROD_INST_ID	{get;set;}
    public	string	宽带账号	{get;set;}
    public	string	关联ITV账号	{get;set;}
    public	string	套餐	{get;set;}
    public	string	支局	{get;set;}
    public	string	网格	{get;set;}
    public	float	CUST_ID	{get;set;}
    public	string	CUST_NAME	{get;set;}
    public	float	AMOUNT	{get;set;}
    public	float	EPON_TYPE	{get;set;}
    public	float	停机类型	{get;set;}
    public	string	宽带速率属性ID	{get;set;}
    public	string	速率	{get;set;}
    public	string	用户联系名称	{get;set;}
    public	float	用户联系方式	{get;set;}
    public	datetime	CRM竣工时间	{get;set;}
    public	string	装机地址	{get;set;}
    public	string	客户地址	{get;set;}
    public	string	F20	{get;set;}
    public	string	F21	{get;set;}

}
