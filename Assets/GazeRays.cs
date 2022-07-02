using System.Runtime.InteropServices;
using UnityEngine;
using ViveSR.anipal.Eye;
using System;
using System.Collections.Generic;
using System.Numerics;
using UnityEngine.UI;
using Vector3 = UnityEngine.Vector3;
using Vector4 = UnityEngine.Vector4;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Text;
using Random = System.Random;

public class GazeRays : MonoBehaviour
{
    private static EyeData eyeData = new EyeData();
    private bool eye_callback_registered = false;

    // Render gaze rays.
    // Press K to toggle.
    // Left is blue, right is red
    private bool renderVisuals = true;
    public Text focusPointDistance;
    public GameObject cube;
    public GameObject depthClus;
    

    // Freeze the gaze visuals so we can inspect them more easily.
    // Press F to toggle.
    private bool isFrozen = false;

    private GameObject LeftVisual;
    private GameObject RightVisual;
    private GameObject DepthVisual;
    private List<float> listFocusFloats = new List<float>();
    private List<float> listCubeDis = new List<float>();
    private float focusDis = -1.0f;
    private float actualDis = -1.0f;
    private bool depthClu = true;
    private float time_to_cold = 0.1f;
    private List<Vector3> eye_mid_point_list = new List<Vector3>();

    void InitLineRenderer(LineRenderer lr)
    {
        lr.startWidth = 0.001f;
        lr.endWidth = 0.001f;
        lr.material = new Material(Shader.Find("Sprites/Default"));
    }

    void Start()
    {
        if (!SRanipal_Eye_Framework.Instance.EnableEye)
        {
            enabled = false;
            return;
        }

        LeftVisual = new GameObject("Left gaze ray visual");
        RightVisual = new GameObject("Right gaze ray visual");
        DepthVisual = new GameObject("the middle line for the Depth");

        LeftVisual.AddComponent<LineRenderer>();
        RightVisual.AddComponent<LineRenderer>();
        DepthVisual.AddComponent<LineRenderer>();
        {
            LineRenderer lr = LeftVisual.GetComponent<LineRenderer>();
            InitLineRenderer(lr);
            lr.startColor = new Color(1, 1, 1, 0);
            lr.endColor = new Color(1, 1, 1, 0);
        }

        {
            LineRenderer lr = RightVisual.GetComponent<LineRenderer>();
            InitLineRenderer(lr);
            lr.startColor = new Color(1, 1, 1, 0);
            lr.endColor = new Color(1, 1, 1, 0);
        }
        {
            LineRenderer mr = DepthVisual.GetComponent<LineRenderer>();
            InitLineRenderer(mr);
            mr.startColor = new Color(1, 1, 1, 0);
            mr.endColor = new Color(1, 1, 1, 0);
            // mr.startColor = Color.yellow;
            // mr.endColor = Color.yellow;
        } 
    }

    void InitEyeData()
    {
        if (
            SRanipal_Eye_Framework.Status != SRanipal_Eye_Framework.FrameworkStatus.WORKING &&
            SRanipal_Eye_Framework.Status != SRanipal_Eye_Framework.FrameworkStatus.NOT_SUPPORT
        ) return;

        if (
            SRanipal_Eye_Framework.Instance.EnableEyeDataCallback == true &&
            eye_callback_registered == false
        ) {
            SRanipal_Eye.WrapperRegisterEyeDataCallback(
                Marshal.GetFunctionPointerForDelegate((SRanipal_Eye.CallbackBasic)EyeCallback)
            );
            eye_callback_registered = true;
        }
        else if (
            SRanipal_Eye_Framework.Instance.EnableEyeDataCallback == false &&
            eye_callback_registered == true
        ) {
            SRanipal_Eye.WrapperUnRegisterEyeDataCallback(
                Marshal.GetFunctionPointerForDelegate((SRanipal_Eye.CallbackBasic)EyeCallback)
            );
            eye_callback_registered = false;
        }
    }

    // Store the results from GetGazeRay for both eyes
    struct RawGazeRays
    {
        public Vector3 leftOrigin;
        public Vector3 leftDir;

        public Vector3 rightOrigin;
        public Vector3 rightDir;

        // Gaze origin and direction are in local coordinates relative to the
        // camera. Here we convert them to absolute coordinates.
        public RawGazeRays Absolute(Transform t)
        {
            var ans = new RawGazeRays();
            ans.leftOrigin = t.TransformPoint(leftOrigin);
            ans.rightOrigin = t.TransformPoint(rightOrigin);
            ans.leftDir = t.TransformDirection(leftDir);
            ans.rightDir = t.TransformDirection(rightDir);
            return ans;
        }
    }

    void GetGazeRays(out RawGazeRays r)
    {
        r = new RawGazeRays();
        if (eye_callback_registered)
        {
            // These return a bool depending whether the gaze ray is available.
            // We can ignore this return value for now.
            SRanipal_Eye.GetGazeRay(GazeIndex.LEFT, out r.leftOrigin, out r.leftDir, eyeData);
            SRanipal_Eye.GetGazeRay(GazeIndex.RIGHT, out r.rightOrigin, out r.rightDir, eyeData);
        }
        else
        {
            SRanipal_Eye.GetGazeRay(GazeIndex.LEFT, out r.leftOrigin, out r.leftDir);
            SRanipal_Eye.GetGazeRay(GazeIndex.RIGHT, out r.rightOrigin, out r.rightDir);
        }

        // Convert from right-handed to left-handed coordinate system.
        // This fixes a bug in the SRanipal Unity package.
        r.leftOrigin.x *= -1;
        r.rightOrigin.x *= -1;
        // Debug.Log("左边的视线方向（RAW）" + r.leftDir.ToString("F4") + "   右边的视线方向(RAW)："  + r.rightDir.ToString("F4") );
    }

    void RenderGazeRays(RawGazeRays gr)
    {
        LineRenderer llr = LeftVisual.GetComponent<LineRenderer>();
        llr.SetPosition(0, gr.leftOrigin);
        llr.SetPosition(1, gr.leftOrigin + gr.leftDir * 20);
        Vector3 left01 = gr.leftOrigin;
        Vector3 left02 = gr.leftOrigin + gr.leftDir * 20;
        Vector3 LeftLine = new Vector3((left02.x - left01.x), (left02.y - left01.y), (left02.z - left01.z));
        LineRenderer rlr = RightVisual.GetComponent<LineRenderer>();
        rlr.SetPosition(0, gr.rightOrigin);
        rlr.SetPosition(1, gr.rightOrigin + gr.rightDir * 20);
        Vector3 right01 = gr.rightOrigin;
        Vector3 right02 = gr.rightOrigin + gr.rightDir * 20;
        Vector3 RightLine = new Vector3((right02.x - right01.x), (right02.y - right01.y), (right02.z - right01.z));
        // Vector3 midLine = Vector3.Cross(LeftLine, RightLine);
        LineRenderer mlr = DepthVisual.GetComponent<LineRenderer>();
        Vector3 pointa, pointb;
        lineToLineSegment(LeftLine, gr.leftOrigin, RightLine, gr.rightOrigin,out pointa,out pointb);
        Vector3 midPoint = (pointa + pointb) / 2;
        float twoPointDis = (pointa - pointb).magnitude;
        float cubeDis = (cube.transform.position - (gr.leftOrigin + gr.rightOrigin) / 2).magnitude;
        float dis = (midPoint - (gr.leftOrigin + gr.rightOrigin) / 2).magnitude;
        time_to_cold += Time.deltaTime;
        if (time_to_cold >= 0.1f)
        {
            eye_mid_point_list.Add(midPoint);
            if (is_gazed(midPoint))
            {
                if (dis >= 1 && dis <= 12 && twoPointDis<=0.2f)
                {
                    listFocusFloats.Add(dis);
                }

                if (cubeDis >= 1 && cubeDis <= 100 && twoPointDis<=0.2f)
                {
                    listCubeDis.Add(cubeDis - 0.5f);
                }
                focusPointDistance.text = dis.ToString() + "真实距离：" + (cube.transform.position - gr.leftOrigin).magnitude.ToString();
                mlr.SetPosition(0,pointa);
                mlr.SetPosition(1,pointb);
            }
            time_to_cold = 0f;
        }
    }

    private bool is_gazed(Vector3 mid)
    {
        if (eye_mid_point_list.Count < 5) return true;
        for (int i = eye_mid_point_list.Count - 1; i >= eye_mid_point_list.Count - 6; i--)
        {
            float d = Vector3.Distance(mid, eye_mid_point_list[i]);
        }
        return true;
    }
    public void lineToLineSegment(Vector3 linea,Vector3 pointInA,Vector3 lineb,Vector3 pointInB,out Vector3 pointA,out Vector3 pointB){
        Vector4 pa = new Vector4(0f, 0f, 0f,0f);
        Vector4 pb = new Vector4(0f, 0f, 0f,0f);
        Vector4 dir1 =new Vector4(0f, 0f, 0f,0f);
        pa.x = pointInA.x;
        pa.y = pointInA.y;
        pa.z = pointInA.z;
        dir1.Set(linea.x,linea.y,linea.z,0);
        pb = pa + dir1;
        Vector4 qa =new Vector4(0f, 0f, 0f,0f); 
        Vector4 qb = new Vector4(0f, 0f, 0f,0f);
        Vector4 dir2 = new Vector4(0f, 0f, 0f,0f);
        qa.Set(pointInB.x,pointInB.y,pointInB.z,0);
        dir2.Set(lineb.x,lineb.y,lineb.z,0);
        qb = qa + dir2;

        Vector4 u = dir1;
        Vector4 v = dir2;
        Vector4 w = pb - qa;

        float a = Vector4.Dot(u, u);
        float b = Vector4.Dot(u, v); 
        float c = Vector4.Dot(v, v);
        float d = Vector4.Dot(u, w);
        float e = Vector4.Dot(v, w);
        float denominator = a * c - b * b;
        float sc, tc;
        if (denominator < 1e-5)          // The lines are almost parallel
        {
            sc = 0.0f;
            tc = (b > c ? d / b : e / c);  // Use the largest denominator
        }
        else
        {
            sc = (b*e - c*d) / denominator;
            tc = (a*e - b*d) / denominator;
        }

        pointA = new Vector4(0, 0, 0, 0);
        pointA =  pb+ sc * u;
        pointB = new Vector3(0, 0, 0);
        pointB = qa + tc * v;
    }

    void SetRenderVisuals(bool value)
    {
        renderVisuals = value;
        // LeftVisual.GetComponent<LineRenderer>().enabled = value;
        // RightVisual.GetComponent<LineRenderer>().enabled = value;
        // DepthVisual.GetComponent<LineRenderer>().enabled = value;
        // LeftVisual.SetActive(value);
        // RightVisual.SetActive(value);
        // DepthVisual.SetActive(value);
        Color a = LeftVisual.GetComponent<LineRenderer>().material.color;
        a.a = 0.0f;
        LeftVisual.GetComponent<LineRenderer>().material.color= a;
        Color b = RightVisual.GetComponent<LineRenderer>().material.color;
        b.a = 0.0f;
        RightVisual.GetComponent<LineRenderer>().material.color= b;
    }
    
    /// <summary>
    /// 写入数据到CSV文件，覆盖形式
    /// </summary>
    /// <param name="csvPath">要写入的字符串表示的CSV文件</param>
    /// <param name="LineDataList">要写入CSV文件的数据，以string[]类型List表示的行集数据</param>
    public void OpCsv(string type)
    {
        String a = "";
        if (type.Equals("focus"))
        {
            for (int i = 0; i < listFocusFloats.Count; i++)
            {
                a +=(listFocusFloats[i].ToString("F4"));
                if (i < listFocusFloats.Count - 1)
                {
                    a += ",";
                }
            } 
        }else if (type.Equals("actual"))
        {
            for (int i = 0; i < listCubeDis.Count; i++)
            {
                a +=(listCubeDis[i].ToString("F4"));
                if (i < listCubeDis.Count - 1)
                {
                    a += ",";
                }
            }  
        }
        Debug.Log(a);
    }

    private void saveExcel(string path, string filename)
    {
        string filepath = path + filename + ".xlsx";
        // 取得文件的信息
        Debug.Log(filepath);
        FileInfo fileInfo = new FileInfo(filepath);

        using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
        {
           ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
           for (int i = 0; i < listCubeDis.Count && i < listFocusFloats.Count; i++)
           {
               Debug.Log(listCubeDis[i]);
               worksheet.Cells[i + 1, 1].Value = listCubeDis[i];
               worksheet.Cells[i + 1, 2].Value = listFocusFloats[i];
           }
           excelPackage.Save();
        }
    }
    private void SaveCSV(string path, string fileName)
    {
        string Folder = Environment.CurrentDirectory + path; // 文件夹路径
        Debug.Log(Folder);
        if (!System.IO.Directory.Exists(Folder)) // 判断文件夹是否存在
            System.IO.Directory.CreateDirectory(Folder); // 创建文件夹
        using (StreamWriter fw = new StreamWriter(fileName, true)) // 以有序字符写入
        {
            foreach (var dis in listFocusFloats)
            {
                fw.Write(dis);
            }
        }
    }

    private void readExcel()
    {
        string filePath = "C:/Users/asd/SteamVRHello/Assets/Result/test.xlsx" ;
        using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets[1];
            Debug.Log(sheet.Cells[1, 1].Value.ToString());
        }
    }
    void Update()
    {
        if (Input.GetKeyDown(KeyCode.K))
        {
            // Toggle visuals
            SetRenderVisuals(!renderVisuals);
        }

        if (Input.GetKeyDown(KeyCode.F))
        {
            // Freeze visuals
            isFrozen = !isFrozen;
        }
        if(Input.GetKeyDown(KeyCode.E))
        {
            Vector3 pos = cube.transform.position;
            Debug.Log(cube.transform.position);
            pos.z = pos.z + 2f;
            cube.transform.position = pos;
        }

        if (Input.GetKeyDown(KeyCode.S))
        {
            Debug.Log("感知深度:");
            OpCsv("focus");
            Debug.Log("实际距离：");
            OpCsv("actual");
            saveExcel(Environment.CurrentDirectory + "/Result/", "result");
        }

        if (Input.GetKeyDown(KeyCode.D))
        {
            depthClus.SetActive(depthClu);
            depthClu = !depthClu;
        }

        if (Input.GetKeyDown(KeyCode.R))
        {
            listCubeDis.Clear();
            listFocusFloats.Clear();
            eye_mid_point_list.Clear();
        }
        if (Input.GetKeyDown(KeyCode.Space))
        {
            float suma = 0, sumb = 0;
            foreach (float number in listFocusFloats)
            {
                suma += number;
            }

            foreach (var dis in listCubeDis)
            {
                sumb += (dis - 0.5f);
            }

            focusDis = suma / listFocusFloats.Count * 1.0f;
            actualDis = sumb / listCubeDis.Count * 1.0f;
            Debug.Log("平均感知深度为：" +focusDis.ToString()  + "  方块平均距离为：" +  actualDis.ToString()+
                      "  感知深度与真实深度的比例为:" + ((focusDis) /(actualDis)).ToString("f4") ); 
        }

        if (Input.GetKeyDown(KeyCode.J))
        {
            if (focusDis > 3 && focusDis <= 12)
            {
                float ratio = focusDis / actualDis;
                Vector3 temp = cube.transform.position;
                temp.z = temp.z / ratio;
                cube.transform.position = temp;
            }
        }
        if (renderVisuals && !isFrozen)
        {
            InitEyeData();

            RawGazeRays localGazeRays;
            GetGazeRays(out localGazeRays);
            RawGazeRays gazeRays = localGazeRays.Absolute(Camera.main.transform);
            // Debug.Log("左边的视线方向：" + gazeRays.leftDir.ToString("F4")  + "  右边的视线方向: " + gazeRays.rightDir.ToString("F4") );
            RenderGazeRays(gazeRays);
        }
    }

    void Release()
    {
        if (eye_callback_registered == true)
        {
            SRanipal_Eye.WrapperUnRegisterEyeDataCallback(
                Marshal.GetFunctionPointerForDelegate((SRanipal_Eye.CallbackBasic)EyeCallback)
            );
            eye_callback_registered = false;
        }

        Destroy(LeftVisual);
        Destroy(RightVisual);
        Destroy(DepthVisual);
    }

    private static void EyeCallback(ref EyeData eye_data)
    {
        eyeData = eye_data;
    }
}
