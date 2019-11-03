using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelDemo.Adapter;
using Microsoft.AspNetCore.Mvc;

namespace ExcelDemo.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            ExcelReader reader = new ExcelReader();
            Console.Write("Read for first time -  "+reader.ReadStream().Count());
            Console.Write("Writting for first time -  " + reader.updateFile().Count());
            Console.Write("Read for second time -  " + reader.ReadStream().Count());
            Console.Write("Writting for second time -  " + reader.updateFile().Count());
            Console.Write("Read for third time -  " + reader.ReadStream().Count());
            Console.Write("Writting for third time -  " + reader.updateFile().Count());
            Console.Write("Read for fourth time -  " + reader.ReadStream().Count());
            Console.Write("Writting for fourth time -  " + reader.updateFile().Count());
            Console.Write("Read for fifth time -  " + reader.ReadStream().Count());
            Console.Write("Writting for fifth time -  " + reader.updateFile().Count());


            return Ok(reader.ReadStream());
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public ActionResult<string> Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
