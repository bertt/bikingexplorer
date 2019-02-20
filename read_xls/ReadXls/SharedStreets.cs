namespace ReadXls
{

    public class Rootobject
    {
        public string type { get; set; }
        public Feature[] features { get; set; }
    }

    public class Feature
    {
        public string type { get; set; }
        public Properties properties { get; set; }
        public Geometry geometry { get; set; }
    }

    public class Properties
    {
        public float score { get; set; }
        public float location { get; set; }
        public float referenceLength { get; set; }
        public string geometryId { get; set; }
        public string referenceId { get; set; }
        public string direction { get; set; }
        public float bearing { get; set; }
        public string snappedSide { get; set; }
        public float interceptAngle { get; set; }
    }

    public class Geometry
    {
        public string type { get; set; }
        public float[] coordinates { get; set; }
    }

}
