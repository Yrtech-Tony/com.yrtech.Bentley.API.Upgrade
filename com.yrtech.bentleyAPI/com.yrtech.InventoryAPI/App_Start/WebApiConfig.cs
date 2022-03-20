using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace com.yrtech.InventoryAPI
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API 配置和服务

            var globalCors = new EnableCorsAttribute("*", "*", "*");
            config.EnableCors(globalCors);

            // Web API 路由
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
            config.Formatters.JsonFormatter.SerializerSettings.Converters.Add(new IsoDateTimeConverter()
            {
                DateTimeFormat = "yyyy-MM-dd HH:mm:ss"
            });
        }
    }
}
