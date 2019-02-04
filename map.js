mapboxgl.accessToken = "pk.eyJ1IjoiYmVydHQiLCJhIjoiaWFvN1Z4USJ9.k04Q6CVms4Gx7KDZxvAx2A";

var map = new mapboxgl.Map({
    container: "map",
    style: "mapbox://styles/mapbox/streets-v9",
    maxZoom: 15,
    minZoom: 10,
    center: [5.121303,52.088551],
    zoom: 11.75,
    bearing: -2.35,
    pitch: 60.0,
    hash: true
});

map.on('load', function() {

map.addLayer({
    'id': 'extrusion',
    'type': 'fill-extrusion',
    "source": {
      "type": "geojson",
      "data": {
        "type": "FeatureCollection",
        "features": []
      }
    },
    'paint': {
        "fill-extrusion-color": {"base": 1,
        "type": "interval",
        "property": "0",
        "default": "#800026",
        "stops": [[0, "#fee8c8"],
                  [5000, "#fdbb84"],
                  [10000, "#fc8d59"],
                  [15000, "#ef6548"],
                  [20000, "#d7301f"],
                  [25000, "#b30000"],
                  [30000, "#7f0000"]]},
      'fill-extrusion-height': ['get', 'traffic'],
      'fill-extrusion-base': 0,
      'fill-extrusion-opacity': 0.9
    }
  });


    map.addSource("meetpunten", {
            "type": "geojson",
            "data": "./utrecht_points.geojson"
    });
    
    map.addLayer({
        "id": "total",
        'type': 'circle',
        'paint': {
          'circle-radius': {
            'base': 5,
            'stops': [
              [12, 5],
              [22, 5]
            ]
          },
          'circle-color': '#ff7770'
        },
        "source": "meetpunten"
      });

    map.on('sourcedata', function(e) {
        if (e.sourceId !== 'meetpunten') return
        if (e.isSourceLoaded !== true) return

        var data1 = {
            "type": "FeatureCollection",
            "features": []
          }
      
          map.querySourceFeatures('meetpunten').forEach(function(f) {
            var object = turf.centerOfMass(f)
            var center = object.geometry.coordinates
            var radius = 50;
            var options = {
              steps: 16,
              units: 'meters',
              properties: object.properties
            };
            data1.features.push(turf.circle(center, radius, options))
          })
          map.getSource('extrusion').setData(data1);        
    });
});
map.getCanvas().style.cursor = "default";






