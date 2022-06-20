using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace SlideSharp
{
    public class Notes
    {
        internal readonly Slides _slides;

        internal Notes(Slides slides)
        {
            this._slides = slides;
        }

        /// <summary>
        /// 添加记录
        /// </summary>
        /// <param name="contents"></param>
        public void Create(string contents = "")
        {
            SlidePart slidePart = _slides._ppt.Slides.SlidePart;
            var relationshipId = _slides._ppt.Slides.RelationshipId;
            NotesSlidePart notesSlidePart;
            string existingSlideNote = "";

            if (slidePart.NotesSlidePart != null)
            {
                var innerText = slidePart.NotesSlidePart.NotesSlide.InnerText;
                existingSlideNote = innerText.Length > 0 ? slidePart.NotesSlidePart.NotesSlide.InnerText + "\n\n" : string.Empty;
                notesSlidePart = slidePart.NotesSlidePart;
            }
            else
            {
                notesSlidePart = slidePart.AddNewPart<NotesSlidePart>(relationshipId);
            }

            NotesSlide notesSlide = new NotesSlide(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = 2U, Name = "" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.SlideImage })),
                        new P.ShapeProperties()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = 3U, Name = "" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1U })),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new BodyProperties(),
                            new ListStyle(),
                            new Paragraph(
                                new Run(
                                    new RunProperties() { Language = "en-US", Dirty = false },
                                    new D.Text() { Text = existingSlideNote + contents }),
                                new EndParagraphRunProperties() { Language = "en-US", Dirty = false }))
                            ))),
                new ColorMapOverride(new MasterColorMapping()));
            notesSlidePart.NotesSlide = notesSlide;
        }
    }
}
