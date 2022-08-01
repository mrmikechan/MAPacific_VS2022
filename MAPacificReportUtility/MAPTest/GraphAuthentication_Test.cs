using Xunit;
using System;
using GraphAuthentication;

namespace MAPTest
{
    public class GraphAuthentication_Test
    {
        [Fact]
        public void Settings_Load()
        {
            GraphAuthentication.Settings settings = GraphAuthentication.Settings.LoadSettings();
            
            
            Assert.Equal(settings.ClientId, "0a4a8270-742e-4711-8362-998ecce2d995");
            Assert.Equal(settings.TenantId, "483e52ff-e285-4763-b1c2-25b904bc8e19");
        }
    }
}