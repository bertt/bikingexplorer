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
        public string Richting1Values { get; set; }
        public string Richting2Values { get; set; }

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

        public override string ToString()
        {
            var result = $"{Id}, {Name}, {LatLon}, r1, {Richting1Dir}, ({Richting1Values}), r2, {Richting2Dir},({Richting2Values})";
            return result;
        }
    }
}
