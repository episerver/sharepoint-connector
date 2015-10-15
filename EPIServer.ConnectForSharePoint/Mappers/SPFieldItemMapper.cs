using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;
using System.Collections.Generic;
using System.Linq;
using System;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPFieldItemMapper : IMapper<JsonFieldItem.D, SPFieldItem>
    {
        public SPFieldItem Map(JsonFieldItem.D source)
        {
            return new SPFieldItem()
            {
                CanBeDeleted = source.CanBeDeleted,
                DefaultValue = source.DefaultValue,
                Description = source.Description,
                Direction = source.Direction,
                EnforceUniqueValues = source.EnforceUniqueValues,
                EntityPropertyName = source.EntityPropertyName,
                Filterable = source.Filterable,
                FromBaseType = source.FromBaseType,
                Group = source.Group,
                Hidden = source.Hidden,
                Id = source.Id,
                Indexed = source.Indexed,
                InternalName = source.InternalName,
                JSLink = source.JSLink,
                ReadOnlyField = source.ReadOnlyField,
                Required = source.Required,
                SchemaXml = source.SchemaXml,
                Scope = source.Scope,
                Sealed = source.Sealed,
                Sortable = source.Sortable,
                StaticName = source.StaticName,
                Title = source.Title, // display name
                FieldTypeKind = source.FieldTypeKind,
                TypeAsString = source.TypeAsString,
                TypeDisplayName = source.TypeDisplayName,
                TypeShortDescription = source.TypeShortDescription,
                ValidationFormula = source.ValidationFormula,
                ValidationMessage = source.ValidationMessage,
                FillInChoice = source.FillInChoice,
                Mappings = source.Mappings,
                Choices = source.Choices,
                EditFormat = source.EditFormat
            };
        }

        public IEnumerable<SPFieldItem> MapList(JsonFieldItem.D[] source)
        {
            return source.Select<JsonFieldItem.D, SPFieldItem>(jsData => Map(jsData));
        }
    }
}