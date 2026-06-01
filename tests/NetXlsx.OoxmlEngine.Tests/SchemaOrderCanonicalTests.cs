// I-82 engine swap — machine-check that OoxmlSchemaOrder's child-order lists match
// the SDK's own compiled schema particle (the order OpenXmlValidator enforces).
//
// OoxmlSchemaOrder hand-maintains the CT_Worksheet / CT_Workbook child-name
// sequences so the engine can insert structural elements at their correct schema
// position on opened files (SDK-quirk #8). A hand-list can silently fall behind the
// schema — and did: the original Type[] form omitted <legacyDrawing>/<legacyDrawingHF>
// (worksheet) and could not even represent <absPath> (workbook), each a latent
// out-of-order-insert bug that no created-from-scratch fixture surfaces. This test
// removes the hand-maintenance risk: it reflects DocumentFormat.OpenXml's compiled
// particle metadata for CT_Worksheet / CT_Workbook, derives the canonical ordered
// list of child local names, and asserts OoxmlSchemaOrder's arrays match it exactly.
// A missing, extra, or misordered name — whether from human error or an SDK upgrade
// that changes the schema — fails the build. The property that matters is
// "an incomplete order list cannot ship," not the reflection mechanism; the
// reflection is intentionally confined to this test so the engine's hot path stays
// free of SDK-internal introspection.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using AwesomeAssertions;
using DocumentFormat.OpenXml;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class SchemaOrderCanonicalTests
{
    [Fact]
    public void WorksheetChildOrder_Matches_The_Sdk_Compiled_Particle()
    {
        string[] canonical = CanonicalChildOrder(new S.Worksheet());
        string[] declared = DeclaredOrder("WorksheetChildOrder");
        declared.Should().Equal(canonical,
            "OoxmlSchemaOrder.WorksheetChildOrder must match the SDK's CT_Worksheet "
            + "child-particle order element-for-element (SDK-quirk #8)");
    }

    [Fact]
    public void WorkbookChildOrder_Matches_The_Sdk_Compiled_Particle()
    {
        string[] canonical = CanonicalChildOrder(new S.Workbook());
        string[] declared = DeclaredOrder("WorkbookChildOrder");
        declared.Should().Equal(canonical,
            "OoxmlSchemaOrder.WorkbookChildOrder must match the SDK's CT_Workbook "
            + "child-particle order element-for-element (SDK-quirk #8)");
    }

    // Reads a private/internal static string[] field off the (internal)
    // OoxmlSchemaOrder type by reflection — IVT is not wired for this assembly, so
    // reflection is how the conformance suite reaches the engine's order lists.
    private static string[] DeclaredOrder(string fieldName)
    {
        var type = typeof(Workbook).Assembly.GetType("NetXlsx.OoxmlSchemaOrder")
            ?? throw new InvalidOperationException("NetXlsx.OoxmlSchemaOrder type not found.");
        var field = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException($"OoxmlSchemaOrder.{fieldName} field not found.");
        return (string[])field.GetValue(null)!;
    }

    // Derives the canonical ordered list of child element local names for a root
    // complex type from the SDK's compiled schema particle. The chain is:
    //   OpenXmlElement.Metadata (ElementMetadata)
    //     -> .Particle (CompiledParticle)
    //       -> .Lookup (IEnumerable<LookupItem>)
    //          each LookupItem: .Type (OpenXmlSchemaType).Name (OpenXmlQualifiedName).Name (string local)
    //                           .Path -> "(Sequence:N)" ordinal
    // The root particle of both CT_Worksheet and CT_Workbook is a flat sequence, so
    // the ordinal is a total order 0..n-1; the test asserts that, which also guards
    // against a future SDK change to a non-flat particle (which would invalidate the
    // flat-rank model OoxmlSchemaOrder relies on).
    private static string[] CanonicalChildOrder(OpenXmlElement root)
    {
        object metadata = GetMetadata(root);
        object particle = GetProp(metadata, "Particle", "ElementMetadata.Particle");
        object lookup = GetProp(particle, "Lookup", "CompiledParticle.Lookup");

        var pairs = new List<(int ordinal, string name)>();
        foreach (var item in (IEnumerable)lookup)
        {
            object schemaType = GetProp(item!, "Type", "LookupItem.Type");
            object qname = GetProp(schemaType, "Name", "OpenXmlSchemaType.Name");
            var localName = (string)GetProp(qname, "Name", "OpenXmlQualifiedName.Name");

            object path = GetProp(item!, "Path", "LookupItem.Path");
            pairs.Add((ParseSequenceOrdinal(path), localName));
        }

        pairs.Should().NotBeEmpty("the SDK must expose a compiled particle for the type");
        pairs.Select(p => p.ordinal).OrderBy(o => o)
            .Should().Equal(Enumerable.Range(0, pairs.Count),
                "the root particle is expected to be a flat 0..n-1 sequence; a gap or "
                + "duplicate means the SDK schema shape changed and the flat-rank model "
                + "in OoxmlSchemaOrder needs revisiting");

        return pairs.OrderBy(p => p.ordinal).Select(p => p.name).ToArray();
    }

    // Path.ToString() is "(Sequence:N)" for a flat-sequence leaf. Parse the trailing
    // integer; a nested particle would not match this shape and throws, surfacing the
    // schema-shape change loudly.
    private static int ParseSequenceOrdinal(object path)
    {
        string text = path.ToString() ?? "";
        int colon = text.LastIndexOf(':');
        int close = colon >= 0 ? text.IndexOf(')', colon) : -1;
        if (colon < 0 || close < 0 || !int.TryParse(text.AsSpan(colon + 1, close - colon - 1), out int ordinal))
            throw new InvalidOperationException(
                $"Unexpected ParticlePath shape '{text}'; the CT_Worksheet/CT_Workbook "
                + "root particle is assumed to be a flat sequence.");
        return ordinal;
    }

    private static object GetMetadata(OpenXmlElement element)
    {
        for (var type = element.GetType(); type is not null; type = type.BaseType)
        {
            var prop = type.GetProperty("Metadata",
                BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.DeclaredOnly);
            if (prop is not null)
                return prop.GetValue(element)
                    ?? throw new InvalidOperationException("OpenXmlElement.Metadata returned null.");
        }
        throw new InvalidOperationException("OpenXmlElement.Metadata property not found via reflection.");
    }

    private static object GetProp(object target, string name, string label)
    {
        var prop = target.GetType().GetProperty(name,
            BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)
            ?? throw new InvalidOperationException($"{label}: property '{name}' not found on {target.GetType().FullName}.");
        return prop.GetValue(target)
            ?? throw new InvalidOperationException($"{label}: property '{name}' returned null.");
    }
}
