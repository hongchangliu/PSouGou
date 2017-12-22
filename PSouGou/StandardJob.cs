using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PSouGou
{
    class StandardJob
    {
        public String 序号 { set; get; }
        public String 城市 { set; get; }
        public String 主点类型 { set; get; }
        public String 补充主点名称 { set; get; }
        public String 现有主点名称 { set; get; }
        public String 现有主点搜狗ID { set; get; }
        public String 别名 { set; get; }
        public String 现有主点是否结构化正确 { set; get; }
        public String 地址 { set; get; }
        public String 大类 { set; get; }
        public String 小类 { set; get; }
        public String 电话 { set; get; }
        public String X { set; get; }
        public String Y { set; get; }
        public String 备注 { set; get; }
        public String 标注员 { set; get; }

        override
        public String ToString() 
        {
            return this.序号 + "," + this.城市 + "," + this.主点类型;
        }
    }
}
