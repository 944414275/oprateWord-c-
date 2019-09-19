using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpireWord.modul
{
    /// <summary>
    /// 数据实体类
    /// </summary>
    public class Student
    {
        //名字
        private string name;
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        //分数
        private string score;
        public string Score
        {
            get { return score; }
            set { score = value; }
        }

        //学生班级
        private string stuClass;
        public string StuClass
        {
            get { return stuClass; }
            set { stuClass = value; }
        }

        //班主任
        private string leader;
        public string Leader
        {
            get { return leader; }
            set { leader = value; }

        }

        public string[] propertyIndex = { "Leader", "Name", "Score", "StuClass" };

        public string GetValue(string PropertyName)
        {
            return (string)this.GetType().GetProperty(PropertyName).GetValue(this, null);
        }


        public List<Student> getStuData()
        {
            List<Student> datas = new List<Student>();
            datas.Add(new Student { Leader = "1", Name = "11", Score = "111", StuClass = "1111" });
            datas.Add(new Student { Leader = "2", Name = "22", Score = "222", StuClass = "2222" });
            datas.Add(new Student { Leader = "3", Name = "33", Score = "333", StuClass = "3333" });
            datas.Add(new Student { Leader = "4", Name = "44", Score = "444", StuClass = "4444" });
            //var cate = datas.GroupBy(s => s.StuClass);
            return datas;
        }
    }
}
