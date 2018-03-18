using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Text.RegularExpressions;

// Class used to represent an AutoCAd text object and facilitate passing
// text object information from the Dwg class methods (that get the information from autoCAD)
// to the refDes class (which prints the information to the xml file).

// Top right cooridnate is lazily loaded so it isn't calculated unless it is needed

public class TextEntity : IDisposable
{
    public String text;
    public DBText dbText;
    public ObjectId id;
    public Point3d BottomLeft;
    //Point3d TopRight;

    //private Lazy<Point3d> _topRightL;
    Point3d _topRightL;

    Regex viewLetterRegex = new Regex(@"^([A-Za-z]){1,2}$");
    Regex calloutRegex = new Regex(@"^(\d)+");
    Regex refDesRegex = new Regex(@"^[A-Z]+(\S)*(\d)+(\s)*(\S)*$");

    public System.Boolean topRightApproximated = false;

    // textEntity Ctor
    public TextEntity()
    {
        text = "";
        dbText = new DBText();
        id = new ObjectId();
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            dbText.Dispose();
        }
    }

    public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }

    public System.Boolean isNull() { return id.IsNull; }

    // textEntity Ctor
    public TextEntity(DBText dbTextIn)
    {
        BottomLeft = dbTextIn.Position;
        text = dbTextIn.TextString;
        dbText = dbTextIn;
        id = dbTextIn.Id;

        // Get vector from bottom left point to center point
        Vector3d vector = dbText.Position.GetVectorTo(dbText.AlignmentPoint);

        // _topRightL = new Lazy<Point3d>(new Func<Point3d>(() => new Point3d(dbText.GeometricExtents.MaxPoint.X, dbText.GeometricExtents.MaxPoint.Y, 0)));

        try
        {
            _topRightL = new Point3d(dbTextIn.GeometricExtents.MaxPoint.X, dbTextIn.GeometricExtents.MaxPoint.Y, 0);
        }
        catch
        {

            double topRightX = Math.Abs(dbText.AlignmentPoint.X + vector.X);
            double topRightY = Math.Abs(dbText.AlignmentPoint.Y + vector.Y);

            if (topRightX <= this.BottomLeft.X || topRightY <= this.BottomLeft.Y)
            {
                // Try using height x width factor x # of chars
                _topRightL = new Point3d(dbText.Position.X * dbText.Height * dbText.TextString.Trim().Length, dbText.Position.Y + dbText.Height, 0);
                topRightApproximated = true;
            }
            else
            {
                _topRightL = new Point3d(topRightX, topRightY, 0);
            }

            //_topRightL = new Point3d (((dbText.AlignmentPoint.X + vector.X > 0) ? dbText.AlignmentPoint.X + vector.X : dbText.Position.X),
            //                                                    ((dbText.AlignmentPoint.Y + vector.Y > 0) ? dbText.AlignmentPoint.Y + vector.Y : dbText.Position.Y),
            //                                                    0);
        }

        //_topRightL = new Lazy<Point3d>(new Func<Point3d>(() => new Point3d(dbText.AlignmentPoint.X + vector.X, dbText.AlignmentPoint.Y + vector.Y, 0)));
        //_topRightL = new Lazy<Point3d>(new Func<Point3d>(() => new Point3d(((dbText.AlignmentPoint.X + vector.X > 0) ? dbText.AlignmentPoint.X + vector.X : dbText.Position.X),
        //  ((dbText.AlignmentPoint.Y + vector.Y > 0) ? dbText.AlignmentPoint.Y + vector.Y : dbText.Position.Y),
        //  0)));
    }

    // textEntity Ctor
    //public textEntity(String textIn, DBText dbTextIn, ObjectId idIn)
    //{
    //    text = textIn;
    //    dbText = dbTextIn;
    //    id = idIn;

    //    // Get vector from bottom left point to center point
    //    Vector3d vector = dbText.Position.GetVectorTo(dbText.AlignmentPoint);

    //    //TopRight = new Point3d(dbText.AlignmentPoint.X + vector.X, dbText.AlignmentPoint.Y + vector.Y, 0);

    //    // Calculate top right coordinate lazily by adding the vector to the center (alignment) point
    //    // Make the top right coordinate the same as the 'position' coordinate (bottom-left) if the alignment point (center coordinate) is missing,
    //    // because the top right coordinate can't be calculated without the alignment point.

    //    _topRightL = new Lazy<Point3d>(new Func<Point3d>(() => new Point3d(dbText.GeometricExtents.MaxPoint.X, dbText.GeometricExtents.MaxPoint.Y, 0)));
    //    // _topRightL = new Lazy<Point3d>(new Func<Point3d>(() => new Point3d(((dbText.AlignmentPoint.X + vector.X > 0) ? dbText.AlignmentPoint.X + vector.X : dbText.Position.X),
    //    //                                                                     ((dbText.AlignmentPoint.Y + vector.Y > 0) ? dbText.AlignmentPoint.Y + vector.Y : dbText.Position.Y),
    //    //                                                                     0)));
    //}

    // Use property to access lazy top right coordinate value
    // public Point3d TopRightL { get { return _topRightL.Value; } }
    public Point3d TopRightL { get { return _topRightL; } }

    // Returns true if this text object is a view letter
    public System.Boolean isViewLetter()
    {
        return (viewLetterRegex.IsMatch(text));
    }

    // Returns true if this text object is a callout
    public System.Boolean isCallout()
    {
        return (calloutRegex.IsMatch(text) && !text.Contains("-"));
    }

    // Returns true if this text object is other text
    public System.Boolean isOther()
    {
        return (!refDesRegex.IsMatch(text) &&
                !this.isOldRefdes() &&
                !this.isCallout() &&
                !this.isViewLetter());
    }

    // Returns true if this text object is an old refdes
    public System.Boolean isOldRefdes()
    {
        return (text.Trim().ToUpper().EndsWith("(REF)"));
    }
}

