using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using MonitorSettings.cs;

namespace KompasProgram
{
    public class KompasProgram
    {
        private MonitorSetting _monitorSetting = new MonitorSetting();
        private KompasObject _kompasObject;
        private ksDocument3D _ksDocument3D;
        private ksPart _ksPart;
        //private ksSketchDefinition _ksSketchDefinition;

        private void RunningTheCompas()
        {
            if (_kompasObject != null)
            {
                _kompasObject.Visible = true;
                _kompasObject.ActivateControllerAPI();
            }
            if (_kompasObject == null)
            {
                Type type = Type.GetTypeFromProgID("KOMPAS.Application.5");
                _kompasObject = (KompasObject)Activator.CreateInstance(type);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wightMonitor"></param>
        /// <param name="heightMonitor"></param>
        /// <param name="heightStande"></param>
        /// <param name="wightStande"></param>
        /// <param name="radius"></param>
        public void SetParametr(MonitorSetting settings)
        {
            _monitorSetting = settings;
        }

        private void CreateANewDocument()
        {
            _ksDocument3D = (ksDocument3D)_kompasObject.Document3D();
            _ksDocument3D.Create();
        }

        private void GetTheComponentInterface()
        {
            const int pTop_Part = -1;
            _ksPart = (ksPart)_ksDocument3D.GetPart(pTop_Part);
        }

        private void CreatSketchThree(ksPart _part, ksEntity ksEntityDrawStande, ksEntity ksEntityPlaneOffset)
        {

            ksSketchDefinition ksSketchDefinition;
            ksDocument2D _ksDocument2D;
            ksEntity ksEntityPlane;
            ksSketchDefinition = (ksSketchDefinition)ksEntityDrawStande.GetDefinition();
            ksEntityPlane = (ksEntity)_part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            ksSketchDefinition.SetPlane(ksEntityPlaneOffset);
            ksEntityDrawStande.Create();
            ksRectangleParam ksRectangleParam;
            ksRectangleParam = (ksRectangleParam)_kompasObject.GetParamStruct((int)StructType2DEnum.ko_RectangleParam);
            ksRectangleParam.ang = 0;
            ksRectangleParam.height = _monitorSetting.StandHeight;
            ksRectangleParam.width = _monitorSetting.StandeWight;
            ksRectangleParam.style = 1;
            ksRectangleParam.x = _monitorSetting.MonitorWigth / 2 - _monitorSetting.StandeWight / 2;
            ksRectangleParam.y = _monitorSetting.MonitorHeight / 2;
            _ksDocument2D = (ksDocument2D)ksSketchDefinition.BeginEdit();
            _ksDocument2D.ksRectangle(ksRectangleParam, 0);
            ksSketchDefinition.EndEdit();
        }

        private void Exstrusion(double size, ksEntity ksEntity, string name,ksEntity ksEntityExtrusion)
        {
            ksEntityExtrusion = (ksEntity)_ksPart.NewEntity((int)Obj3dType.o3d_baseExtrusion);
            ksBaseExtrusionDefinition ksBaseExtrusionDefinition = (ksBaseExtrusionDefinition)ksEntityExtrusion.GetDefinition();
            ksBaseExtrusionDefinition.SetSideParam(true, (short)End_Type.etBlind, size, 0, true);
            ksBaseExtrusionDefinition.SetSketch(ksEntity);
            ksEntityExtrusion.name = name;
            ksEntityExtrusion.useColor = 0;
            ksEntityExtrusion.SetAdvancedColor(3223857, 0.9, 0.8, 0.7, 0.6, 1, 0.4);
            ksEntityExtrusion.Create();
        }
        private void CreatSketch(ksPart _part, ksEntity ksEntityDraw)
        {
            ksSketchDefinition ksSketchDefinition;
            ksDocument2D _ksDocument2D;
            ksEntity ksEntityPlane;
            ksSketchDefinition = (ksSketchDefinition)ksEntityDraw.GetDefinition();
            ksEntityPlane = (ksEntity)_part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            ksSketchDefinition.SetPlane(ksEntityPlane);
            ksEntityDraw.Create();
            ksRectangleParam ksRectangleParam;
            ksRectangleParam = (ksRectangleParam)_kompasObject.GetParamStruct((int)StructType2DEnum.ko_RectangleParam);
            ksRectangleParam.ang = 0;
            ksRectangleParam.height = _monitorSetting.MonitorHeight + 70;
            ksRectangleParam.width = _monitorSetting.MonitorWigth + 50;
            ksRectangleParam.style = 1;
            ksRectangleParam.x = 0;
            ksRectangleParam.y = 0;
            _ksDocument2D = (ksDocument2D)ksSketchDefinition.BeginEdit();
            _ksDocument2D.ksRectangle(ksRectangleParam, 0);
            ksSketchDefinition.EndEdit();
        }
        private void CreatSketchCircle(ksPart _part, double wight, double heigth,
            ksEntity ksEntityCircle, ksEntity _ksEntityPlaneOffset)
        {
            ksSketchDefinition ksSketchDefinition;
            ksDocument2D _ksDocument2D;
            ksEntity ksEntityPlane;
            ksSketchDefinition = (ksSketchDefinition)ksEntityCircle.GetDefinition();
            ksEntityPlane = (ksEntity)_part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            ksSketchDefinition.SetPlane(_ksEntityPlaneOffset);
            ksEntityCircle.Create();
            _ksDocument2D = (ksDocument2D)ksSketchDefinition.BeginEdit();
            _ksDocument2D.ksCircle(_monitorSetting.MonitorWigth / 2,
               -90, _monitorSetting.Radius, 1);
            ksSketchDefinition.EndEdit();
        }

        private void CreatSketchTwo(ksPart _part, ksEntity ksEntityDrawTwo)
        {
            ksSketchDefinition _ksSketchDefinition;
            ksDocument2D _ksDocument2D;
            ksEntity ksEntityPlane;
            _ksSketchDefinition = (ksSketchDefinition)ksEntityDrawTwo.GetDefinition();
            ksEntityPlane = (ksEntity)_part.GetDefaultEntity((short)Obj3dType.o3d_planeXOY);
            _ksSketchDefinition.SetPlane(ksEntityPlane);
            ksEntityDrawTwo.Create();
            ksRectangleParam ksRectangleParam;
            ksRectangleParam = (ksRectangleParam)_kompasObject.GetParamStruct((int)StructType2DEnum.ko_RectangleParam);
            ksRectangleParam.ang = 0;
            ksRectangleParam.height = _monitorSetting.MonitorHeight;
            ksRectangleParam.width = _monitorSetting.MonitorWigth;
            ksRectangleParam.style = 1;
            ksRectangleParam.x = 25;
            ksRectangleParam.y = 35;
            _ksDocument2D = (ksDocument2D)_ksSketchDefinition.BeginEdit();
            _ksDocument2D.ksRectangle(ksRectangleParam, 0);
            _ksSketchDefinition.EndEdit();
        }
        private void ExstrusionCut(string name, ksEntity ksEntityExtrusion, ksEntity _ksEntityDrawTwo)
        {
            ksCutExtrusionDefinition _ksCutExtrusionDefinition = (ksCutExtrusionDefinition)ksEntityExtrusion.GetDefinition();
            _ksCutExtrusionDefinition.SetSideParam(true, (short)End_Type.etBlind, 10, 0, true);
            _ksCutExtrusionDefinition.SetSketch(_ksEntityDrawTwo);
            ksEntityExtrusion.name = name;
            ksEntityExtrusion.useColor = 0;
            ksEntityExtrusion.SetAdvancedColor(0, 0.9, 0.8, 0.7, 0.6, 1, 0.4);
            ksEntityExtrusion.Create();
        }

        private void CreatPlaneOffset(double offset, short plane, ksEntity ksEntityPlaneOffset)
        {
            ksEntity ksEntityPlaneXOY = (ksEntity)_ksPart.GetDefaultEntity(plane);
            ksPlaneOffsetDefinition ksPlaneOffsetDefinition = (ksPlaneOffsetDefinition)ksEntityPlaneOffset.GetDefinition();
            ksPlaneOffsetDefinition.SetPlane(ksEntityPlaneXOY);
            ksPlaneOffsetDefinition.direction = true;
            ksPlaneOffsetDefinition.offset = offset;
            ksEntityPlaneOffset.Create();
        }
        public void Construct()
        {
            RunningTheCompas();
            CreateANewDocument();
            GetTheComponentInterface();
            ksEntity ksEntityDraw = (ksEntity)_ksPart.NewEntity((short)Obj3dType.o3d_sketch);
            ksEntity ksEntityDrawTwo = (ksEntity)_ksPart.NewEntity((short)Obj3dType.o3d_sketch);
            ksEntity ksEntityPlaneOffset = (ksEntity)_ksPart.NewEntity(
                (short)Obj3dType.o3d_planeOffset);
            ksEntity ksEntityExtrusion = (ksEntity)_ksPart.NewEntity(
                (short)Obj3dType.o3d_cutExtrusion);
            ksEntity ksEntityCircle = (ksEntity)_ksPart.NewEntity((short)Obj3dType.o3d_sketch);
            ksEntity ksEntityDrawStande = (ksEntity)_ksPart.NewEntity((short)Obj3dType.o3d_sketch);
            ksEntity ksEntityPlaneOffsetTwo = (ksEntity)_ksPart.NewEntity(
                (short)Obj3dType.o3d_planeOffset);     
            CreatSketch(_ksPart, ksEntityDraw);
            Exstrusion(50, ksEntityDraw, "Выдавливание 1",ksEntityExtrusion);
            CreatSketchTwo(_ksPart, ksEntityDrawTwo);
            ExstrusionCut("Выдавливание вырезанием", ksEntityExtrusion,
                ksEntityDrawTwo);
            CreatPlaneOffset(50, (short)Obj3dType.o3d_planeXOY, ksEntityPlaneOffsetTwo);
            CreatSketchThree(_ksPart, ksEntityDrawStande, ksEntityPlaneOffsetTwo);
            Exstrusion(70, ksEntityDrawStande, "Выдавливание 2",ksEntityExtrusion);
            CreatPlaneOffset(_monitorSetting.MonitorHeight / 2 +_monitorSetting.StandHeight,
                (short)Obj3dType.o3d_planeXOZ, ksEntityPlaneOffset);
            CreatSketchCircle(_ksPart, 20, 20, ksEntityCircle, ksEntityPlaneOffset);
            Exstrusion(20, ksEntityCircle, "Выдавливание 3",ksEntityExtrusion);
            _kompasObject.Visible = true;
        }
    }
}
