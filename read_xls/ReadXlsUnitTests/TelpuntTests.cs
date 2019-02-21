using NUnit.Framework;
using ReadXls;

namespace Tests
{
    public class TelpuntTests
    {
        [Test]
        public void Test1()
        {
            var telpunt1 = new Telpunt() { Richting1 = "Noordoost", Richting2 = "Zuidwest", Forward = 71.83128f, Backward = 251.831284f };
            telpunt1.CheckDirections();
            Assert.IsTrue(telpunt1.Richting1Dir == telpunt1.Forward);
            Assert.IsTrue(telpunt1.Richting2Dir == telpunt1.Backward);


            var telpunt2 = new Telpunt() { Richting1 = "Oost", Richting2 = "West", Forward = 277.798f, Backward = 97.79801f };
            telpunt2.CheckDirections();
            Assert.IsTrue(telpunt2.Richting1Dir == telpunt2.Backward);
            Assert.IsTrue(telpunt2.Richting2Dir == telpunt2.Forward);


            var telpunt3 = new Telpunt() { Richting1 = "Noord", Richting2 = "Zuid", Forward = 350.93396f, Backward = 170.93396f };
            telpunt3.CheckDirections();
            Assert.IsTrue(telpunt3.Richting1Dir == telpunt3.Forward);
            Assert.IsTrue(telpunt3.Richting2Dir == telpunt3.Backward);


            var telpunt4 = new Telpunt() { Richting1 = "Zuid", Richting2 = "Noord", Forward = 125.104134f, Backward = 344.874359f };
            telpunt4.CheckDirections();
            Assert.IsTrue(telpunt4.Richting1Dir == telpunt4.Forward);
            Assert.IsTrue(telpunt4.Richting2Dir == telpunt4.Backward);

            var telpunt5 = new Telpunt() { Richting1 = "West", Richting2 = "Oost", Forward = 93.4628143f, Backward = 272.031677f };
            telpunt5.CheckDirections();
            Assert.IsTrue(telpunt5.Richting1Dir == telpunt5.Backward);
            Assert.IsTrue(telpunt5.Richting2Dir == telpunt5.Forward);

        }
    }
}