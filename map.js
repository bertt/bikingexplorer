mapboxgl.accessToken = "pk.eyJ1IjoiYmVydHQiLCJhIjoiaWFvN1Z4USJ9.k04Q6CVms4Gx7KDZxvAx2A";

var map = new mapboxgl.Map({
    container: "map",
    style: "mapbox://styles/mapbox/streets-v9",
    maxZoom: 15,
    minZoom: 10,
    center: [5.121303,52.088551],
    zoom: 11.75,
    bearing: -2.35,
    pitch: 60.0
});

//center: start_story.center,
//zoom: start_story.zoom,
//bearing: start_story.bearing,
//pitch: start_story.pitch
// alert('hallo');
  
//map.flyTo(start_extent);

map.on("load", function(e) {
    map.addSource("meetpunten", {
            "type": "geojson",
            "data": "./utrecht_points.geojson"
    });
    
    map.addLayer({
        'id': 'meetpunten',
        'source': 'meetpunten',
        'type': 'circle',
        "paint": {
            "circle-color": "#FF0000",
            "circle-radius": 5
        }
    })
});

map.getCanvas().style.cursor = "default";






