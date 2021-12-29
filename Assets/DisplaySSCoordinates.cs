using UnityEngine;
using UnityEngine.UI;
using UnityEngine.XR;

public class DisplaySSCoordinates : MonoBehaviour
{
    public Transform debugSphere;
    public Text displayText;
    public Camera vrCamera;
    
    private void Update()
    {
        var coords = vrCamera.WorldToScreenPoint(debugSphere.position, Camera.MonoOrStereoscopicEye.Left);
        var eyeTextureSize = new Vector2(XRSettings.eyeTextureWidth, XRSettings.eyeTextureHeight);
        var units = new Vector2(coords.x / eyeTextureSize.x, coords.y / eyeTextureSize.y);
        displayText.text = $"({units.x:F2},{units.y:F2})";
    }
}