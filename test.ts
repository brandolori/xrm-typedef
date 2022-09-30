const getPropertietor = <Type>(obj: Type) => {

    return function getProperty<Key extends keyof Type>(key: Key) {
        return obj[key];
    }
}


const a = { first: 10, second: "20" }

const propertietor = getPropertietor(a)

const b = propertietor("first")

const formLoad = (formContext: FormContext<typeof a>) => {
    const a = formContext.getAttribute("second").getValue()
}