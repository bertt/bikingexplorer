using System.Collections.Generic;

namespace ReadXls
{
    public class Telpunt
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string LatLon { get; set; }
        public string Richting1 { get; set; }
        public string Richting2 { get; set; }
        public float Forward { get; set; }
        public float Backward { get; set; }
        public float Richting1Dir { get; set; }
        public float Richting2Dir { get; set; }
        public List<int> Richting1Measurements { get; set; }
        public List<int> Richting2Measurements { get; set; }

        public void CheckDirections()
        {
            Richting1Dir = Forward;
            Richting2Dir = Backward;

            if ((Richting1.ToLower().Contains("oost") && Forward >= 180) || (Richting1.ToLower().Contains("west") && Forward <= 180))
            {
                Richting1Dir = Backward;
                Richting2Dir = Forward;
            }
        }
    }
}
