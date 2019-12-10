using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelReport.Model
{
    public class Response<T>
    {
        public Response() : this(false, "") { }

        public Response(bool success) : this(success, "")
        {
            this.Success = success;
        }

        public Response(bool success, string message)
        {
            Success = success;
            MessageText = message;
            Items = default(T);
        }

        public bool Success { get; set; }

        public string MessageText { get; set; }

        public T Items { get; set; }

   
        /// <summary>
        /// 每页显示数量
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// 总数量
        /// </summary>
        public int TotalCount { get; set; }

       

        /// <summary>
        /// 状态码
        /// </summary>
        public string Status { get; set; }
        public Response<T> Fail()
        {
            this.Success = false;

            return this;
        }

        /// <summary>
        /// 失败
        /// </summary>
        /// <param name="errMessage"></param>
        /// <returns></returns>
        public Response<T> Fail(string errMessage)
        {
            this.MessageText = errMessage;

            this.Success = false;

            return this;
        }

       

        /// <summary>
        /// 返回成功的
        /// </summary>
        /// <returns></returns>
        public Response<T> OK()
        {
            this.Success = true;

            return this;
        }

        public Response<T> OK(T value)
        {
            this.Success = true;
            this.Items = value;
            return this;
        }

        public Response<T> SetTotal(int total)
        {
            this.TotalCount = total;
            return this;
        }
    }

    public class Response
    {
        public bool Success { get; set; }

        public string MessageText { get; set; }

        public object Result { get; set; }

    }
}
