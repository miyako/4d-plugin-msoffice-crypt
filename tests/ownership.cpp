#include "4DPlugin-JSON.h"
#include "PrivateTypes.h"
#include "EntryPoints.h"

#include <cassert>
#include <cstdlib>
#include <cstring>
#include <map>
#include <set>

namespace {
using Text = CUTF16String;
struct Object {
    size_t refs = 1;
    std::map<Text, PA_Variable> properties;
};
struct Collection {
    size_t refs = 1;
};
std::set<PA_Unichar*> strings;
std::map<PA_Handle, size_t> handles;
std::set<Object*> objects;
std::set<Collection*> collections;
bool allocatedEmpty = true;
bool stringifyMismatch = false;

Text text(const PA_Unistring& value)
{
    PA_Unichar* buffer = value.fString;
    assert(value.fLength >= 0);
    if (!value.fString) {
        assert(value.fLength == 0);
        return {};
    }
    assert(strings.count(buffer));
    assert(value.fString[value.fLength] == 0);
    return Text(value.fString, value.fLength);
}

PA_Unistring makeString(const Text& value)
{
    PA_Unistring result = {};
    result.fLength = static_cast<PA_long32>(value.size());
    if (!value.empty() || allocatedEmpty) {
        result.fString = static_cast<PA_Unichar*>(
            std::malloc((value.size() + 1) * sizeof(PA_Unichar)));
        assert(result.fString);
        std::memcpy(result.fString, value.c_str(), (value.size() + 1) * sizeof(PA_Unichar));
        PA_Unichar* buffer = result.fString;
        assert(strings.insert(buffer).second);
    }
    return result;
}

void freeString(PA_Unistring& value)
{
    if (value.fString) {
        PA_Unichar* buffer = value.fString;
        assert(strings.erase(buffer) == 1);
        std::free(value.fString);
    }
    value = {};
}

void clear(PA_Variable& value);

void release(Object* object)
{
    assert(objects.count(object) && object->refs);
    if (--object->refs == 0) {
        for (auto& entry : object->properties)
            clear(entry.second);
        objects.erase(object);
        delete object;
    }
}

void clear(PA_Variable& value)
{
    switch (value.fType) {
        case eVK_Unistring:
            freeString(value.uValue.fString);
            break;
        case eVK_Object:
            if (value.uValue.fObject)
                release(static_cast<Object*>(value.uValue.fObject));
            break;
        case eVK_Collection: {
            auto collection = static_cast<Collection*>(value.uValue.fCollection);
            if (collection) {
                assert(collections.count(collection) && collection->refs);
                if (--collection->refs == 0) {
                    collections.erase(collection);
                    delete collection;
                }
            }
            break;
        }
        case eVK_ArrayUnicode: {
            PA_Handle handle = value.uValue.fArray.fData;
            assert(handles.count(handle));
            assert(handles.at(handle) ==
                   (value.uValue.fArray.fNbElements + 1) * sizeof(PA_Unistring));
            auto elements = reinterpret_cast<PA_Unistring*>(*handle);
            for (PA_long32 i = 0; i <= value.uValue.fArray.fNbElements; ++i)
                freeString(elements[i]);
            std::free(*handle);
            std::free(handle);
            handles.erase(handle);
            break;
        }
        default:
            assert(value.fType == eVK_Undefined || value.fType == eVK_Boolean ||
                   value.fType == eVK_Real || value.fType == eVK_Longint);
    }
    value = {};
    value.fType = eVK_Undefined;
}

PA_Variable copy(PA_Variable value)
{
    switch (value.fType) {
        case eVK_Unistring:
            value.uValue.fString = makeString(text(value.uValue.fString));
            break;
        case eVK_Object:
            if (value.uValue.fObject)
                ++static_cast<Object*>(value.uValue.fObject)->refs;
            break;
        case eVK_Collection:
            if (value.uValue.fCollection)
                ++static_cast<Collection*>(value.uValue.fCollection)->refs;
            break;
    }
    return value;
}

Object* newObject()
{
    auto object = new Object;
    assert(objects.insert(object).second);
    return object;
}

void engine(short selector, EngineBlock* eb)
{
    switch (selector) {
        case EX_CREATE_UNISTRING: {
            auto source = reinterpret_cast<PA_Unichar*>(eb->fHandle);
            eb->fUniString1 = makeString(source ? Text(source) : Text());
            break;
        }
        case EX_DISPOSE_UNISTRING:
            freeString(eb->fUniString1);
            break;
        case EX_SET_UNISTRING: {
            Text source(reinterpret_cast<PA_Unichar*>(eb->fHandle));
            freeString(eb->fUniString1);
            eb->fUniString1 = makeString(source);
            break;
        }
        case EX_VARIABLE_TO_STRING: {
            auto value = reinterpret_cast<PA_Variable*>(eb->fHandle);
            assert(value->fType == eVK_Unistring);
            text(value->uValue.fString);
            eb->fUniString1 = value->uValue.fString;
            break;
        }
        case EX_CLEAR_VARIABLE:
            clear(*static_cast<PA_Variable*>(eb->fPtr1));
            break;
        case EX_GET_INFORMATION:
            eb->fParam2 = 0x00001700;
            break;
        case EX_HANDLE_MANAGER: {
            PA_Handle handle = eb->fHandle;
            switch (eb->fParam1) {
                case 1:
                    handle = static_cast<PA_Handle>(std::malloc(sizeof(char*)));
                    assert(handle);
                    *handle = static_cast<char*>(std::calloc(1, eb->fParam2 ? eb->fParam2 : 1));
                    assert(*handle);
                    handles.emplace(handle, static_cast<size_t>(eb->fParam2));
                    eb->fHandle = handle;
                    break;
                case 2:
                    assert(handles.erase(handle) == 1);
                    std::free(*handle);
                    std::free(handle);
                    break;
                case 3:
                    assert(handles.count(handle));
                    *handle = static_cast<char*>(std::realloc(*handle, eb->fParam2));
                    assert(*handle);
                    handles.at(handle) = eb->fParam2;
                    break;
                case 4:
                    eb->fParam2 = handles.at(handle);
                    break;
                case 5:
                    assert(handles.count(handle));
                    eb->fParam3 = reinterpret_cast<sLONG_PTR>(*handle);
                    break;
                case 6:
                    assert(handles.count(handle));
                    break;
                default:
                    std::abort();
            }
            break;
        }
        case EX_GET_OBJ_VALUE: {
            assert(eb->fShort == eVK_Object);
            auto object = static_cast<Object*>(eb->fPtr1);
            assert(objects.count(object));
            *static_cast<PA_Variable*>(eb->fPtr2) =
                copy(object->properties.at(text(eb->fUniString1)));
            break;
        }
        case EX_SET_OBJ_VALUE: {
            assert(eb->fShort == eVK_Object);
            auto object = static_cast<Object*>(eb->fPtr1);
            assert(objects.count(object));
            Text key = text(eb->fUniString1);
            PA_Variable replacement = copy(*static_cast<PA_Variable*>(eb->fPtr2));
            auto old = object->properties.find(key);
            if (old != object->properties.end())
                clear(old->second);
            object->properties[key] = replacement;
            break;
        }
        case EX_EXECUTE_COMMAND_BY_ID: {
            auto args = static_cast<PA_Variable**>(eb->fPtr1);
            auto result = static_cast<PA_Variable*>(eb->fPtr2);
            switch (eb->fParam1) {
                case 1471:
                    PA_SetObjectVariable(result, newObject());
                    break;
                case 1472: {
                    auto collection = new Collection;
                    collections.insert(collection);
                    PA_SetCollectionVariable(result, collection);
                    break;
                }
                case 1231: {
                    assert(eb->fParam2 == 2);
                    auto object = static_cast<Object*>(args[0]->uValue.fObject);
                    PA_SetBooleanVariable(result,
                        object->properties.count(text(args[1]->uValue.fString)) != 0);
                    break;
                }
                case 1225: {
                    auto source = static_cast<Object*>(args[0]->uValue.fObject);
                    auto duplicate = newObject();
                    for (auto& entry : source->properties)
                        duplicate->properties.emplace(entry.first, copy(entry.second));
                    PA_SetObjectVariable(result, duplicate);
                    break;
                }
                case 1217:
                    assert(eb->fParam2 == 1);
                    if (stringifyMismatch) {
                        auto collection = new Collection;
                        collections.insert(collection);
                        PA_SetCollectionVariable(result, collection);
                    } else {
                        auto value = makeString(Text{'{', '}'});
                        PA_SetStringVariable(result, &value);
                    }
                    break;
                default:
                    std::abort();
            }
            break;
        }
        default:
            std::abort();
    }
}

void balanced()
{
    if (!strings.empty() || !handles.empty() || !objects.empty() || !collections.empty()) {
        std::fprintf(stderr, "ownership imbalance: strings=%zu handles=%zu objects=%zu collections=%zu\n",
                     strings.size(), handles.size(), objects.size(), collections.size());
        std::abort();
    }
}

void setters()
{
    const wchar_t* wide[] = {L"ASCII", L"\u65e5\u672c", L"\U0001f600", L""};
    const char* utf8[] = {"ASCII", "\xe6\x97\xa5\xe6\x9c\xac", "\xf0\x9f\x98\x80", ""};
    const Text expected[] = {{'A','S','C','I','I'}, {0x65e5,0x672c}, {0xd83d,0xde00}, {}};
    for (size_t i = 0; i < 4; ++i) {
        for (int overload = 0; overload < 4; ++overload) {
            PA_ObjectRef object = PA_CreateObject();
            for (int repeat = 0; repeat < 3; ++repeat) {
                switch (overload) {
                    case 0: ob_set_s(object, "key", utf8[i]); break;
                    case 1: ob_set_s(object, L"key", utf8[i]); break;
                    case 2: ob_set_a(object, L"key", wide[i]); break;
                    case 3: ob_set_a(object, L"key", &expected[i]); break;
                }
                const auto& value = static_cast<Object*>(object)->properties.at(Text{'k','e','y'});
                assert(value.fType == eVK_Unistring);
                assert(text(value.uValue.fString) == expected[i]);
                std::vector<Text> churn(128, Text(128, 'x'));
                assert(text(value.uValue.fString) == expected[i]);
            }
            PA_DisposeObject(object);
            balanced();
        }
    }
    ob_set_s(nullptr, "key", "value");
    ob_set_s(nullptr, L"key", "value");
    ob_set_a(nullptr, L"key", L"value");
    balanced();
}

void getters()
{
    PA_ObjectRef object = PA_CreateObject();
    const size_t initialStrings = strings.size();
    ob_set_s(object, L"text", "A\xe6\x97\xa5\xf0\x9f\x98\x80");
    ob_set_s(object, L"empty", "");
    ob_set_b(object, L"bool", true);
    ob_set_n(object, L"number", 3.25);
    PA_CollectionRef collection = PA_CreateCollection();
    ob_set_c(object, L"collection", collection);
    PA_ObjectRef child = PA_CreateObject();
    ob_set_o(object, L"object", child);
    PA_ObjectRef other = PA_CreateObject();
    ob_set_o(other, "object", PA_DuplicateObject(child));
    PA_DisposeObject(other);
    const size_t residentStrings = strings.size();
    for (int repeat = 0; repeat < 20; ++repeat) {
        CUTF8String output;
        assert(ob_get_a(object, L"text", &output));
        assert(output == CUTF8String(reinterpret_cast<const uint8_t*>("A\xe6\x97\xa5\xf0\x9f\x98\x80")));
        assert(ob_get_a(object, L"empty", &output) && output.empty());
        assert(!ob_get_a(object, L"missing", &output));
        assert(ob_get_b(object, L"bool"));
        assert(ob_get_n(object, L"number") == 3.25);
        for (const wchar_t* key : {L"text", L"collection", L"object"}) {
            assert(!ob_get_b(object, key));
            assert(ob_get_n(object, key) == 0);
            assert(ob_get_c(object, key) == (key == std::wstring(L"collection") ? collection : nullptr));
            if (key == std::wstring(L"collection"))
                PA_DisposeCollection(collection);
        }
        assert(ob_get_a(object, L"collection", &output));
        assert(ob_get_a(object, L"object", &output));
        assert(ob_get_a(object, L"bool", &output));
        if (strings.size() != residentStrings) {
            std::fprintf(stderr, "ownership imbalance: property reads retained %zu extra strings\n",
                         strings.size() - residentStrings);
            std::abort();
        }
        assert(static_cast<Object*>(child)->refs == 1);
        assert(static_cast<Collection*>(collection)->refs == 1);
    }
    CUTF8String independent;
    assert(ob_get_a(object, L"text", &independent));
    ob_set_s(object, L"text", "replacement");
    assert(independent == CUTF8String(reinterpret_cast<const uint8_t*>("A\xe6\x97\xa5\xf0\x9f\x98\x80")));
    CUTF8String json;
    ob_stringify(object, &json);
    assert(json == CUTF8String(reinterpret_cast<const uint8_t*>("{}")));
    stringifyMismatch = true;
    ob_stringify(object, &json);
    stringifyMismatch = false;
    PA_CollectionRef transferred = ob_get_c(object, L"collection");
    PA_DisposeObject(object);
    assert(collections.count(static_cast<Collection*>(transferred)));
    PA_DisposeCollection(transferred);
    assert(strings.size() == initialStrings);
    balanced();
}

void arrays()
{
    PA_Variable array = PA_CreateVariable(eVK_ArrayUnicode);
    PA_ResizeArray(&array, 2);
    Text expected{'A', 0x65e5, 0xd83d, 0xde00};
    PA_Unistring input = PA_CreateUnistring(const_cast<PA_Unichar*>(expected.c_str()));
    PA_SetStringInArray(array, 0, &input);
    PA_Unistring borrowed = PA_GetStringInArray(array, 0);
    assert(input.fString == borrowed.fString);
    PA_Unistring independent = PA_CreateUnistring(borrowed.fString);
    assert(independent.fString != borrowed.fString);
    PA_SetStringInArray(array, 1, &independent);
    PA_Unistring replacement = PA_CreateUnistring(const_cast<PA_Unichar*>(Text{'B'}.c_str()));
    PA_SetStringInArray(array, 0, &replacement);
    assert(text(PA_GetStringInArray(array, 1)) == expected);
    PA_ResizeArray(&array, 1);
    assert(text(PA_GetStringInArray(array, 1)) == expected);
    PA_Variable wrong = PA_CreateVariable(eVK_Longint);
    PA_Variable nullData = {};
    nullData.fType = eVK_ArrayUnicode;
    for (int destination = 0; destination < 4; ++destination) {
        auto value = PA_CreateUnistring(const_cast<PA_Unichar*>(expected.c_str()));
        PA_SetStringInArray(destination == 0 ? wrong : destination == 1 ? nullData : array,
                            destination == 2 ? -1 : destination == 3 ? 2 : 0, &value);
        assert(text(value) == expected);
        PA_DisposeUnistring(&value);
    }
    PA_SetStringInArray(array, 0, nullptr);
    PA_ResizeArray(&array, 0);
    PA_ClearVariable(&array);
    balanced();
}
}

extern "C" void PluginMain(PA_long32, PA_PluginParameters) {}

int main(int argc, char** argv)
{
    assert(argc == 2);
    gCall4D = engine;
    for (bool mode : {false, true}) {
        allocatedEmpty = mode;
        if (std::strcmp(argv[1], "setters") == 0)
            setters();
        else if (std::strcmp(argv[1], "getters") == 0)
            getters();
        else if (std::strcmp(argv[1], "arrays") == 0)
            arrays();
        else
            std::abort();
    }
    std::printf("%s: PASS (null and allocated empties)\n", argv[1]);
}
